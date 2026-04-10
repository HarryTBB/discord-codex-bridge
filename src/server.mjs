import fs from "node:fs";
import path from "node:path";

import dotenv from "dotenv";
import { Client, EmbedBuilder, Events, GatewayIntentBits } from "discord.js";

import {
  buildHelpMessage,
  buildInitialScopeState,
  buildProjectScopeState,
  buildStatusMessage,
  chunkDiscordMessage,
  cloneUserProjects,
  createDirectoryWithinProject,
  createProjectDirectory,
  createScopeKey,
  createStateStore,
  findLatestProjectSession,
  getEffectiveUserSettings,
  getAllowedProjectsForUser,
  hasAssignedProjects,
  isAdminUser,
  isAllowedUser,
  listPathLikeLs,
  listTopLevelProjects,
  loadConfig,
  log,
  normalizeDiscordUserId,
  normalizePermissionLevel,
  normalizeProjectName,
  normalizeReasoningValue,
  prepareAttachmentsForPrompt,
  readLatestRateLimits,
  reconcileScopeState,
  resolveScopedProjectPath,
  runCodex,
  serializeKeyValueMap,
  serializeUserProjects,
  switchProjectDirectory,
  validateRuntimeConfig
} from "./shared.mjs";

const envFilePath = path.join(process.cwd(), ".env");
dotenv.config({ path: envFilePath });

let config = loadConfig();
validateRuntimeConfig(config);

const client = new Client({
  intents: [
    GatewayIntentBits.Guilds,
    GatewayIntentBits.DirectMessages,
    ...(config.discordAutoAddOnGuildJoin ? [GatewayIntentBits.GuildMembers] : [])
  ]
});

const state = createStateStore(config.stateDir, config.defaultCwd);
const inflightByScope = new Map();
const adminCommandNames = new Set([
  "admin-user-add",
  "admin-user-remove",
  "admin-user-grant-project",
  "admin-user-set-permission"
]);
const projectRequiredCommandNames = new Set(["ls", "ask", "resume", "project", "project-switch", "mkdir", "pwd", "cd", "reset"]);

client.once(Events.ClientReady, (readyClient) => {
  log("info", "Discord Codex bridge connected", {
    botTag: readyClient.user.tag,
    workspaceRoot: config.workspaceRoot,
    defaultCwd: config.defaultCwd,
    autoAddOnGuildJoin: config.discordAutoAddOnGuildJoin,
    autoAddProjects: config.discordAutoAddProjects,
    autoAddPermission: config.discordAutoAddPermission
  });
});

client.on(Events.GuildMemberAdd, async (member) => {
  if (!config.discordAutoAddOnGuildJoin) {
    return;
  }
  if (member.user.bot) {
    return;
  }
  if (member.guild.id !== config.discordGuildId) {
    return;
  }

  try {
    const outcome = autoOnboardGuildMember(member.user.id);
    log("info", "Guild member auto-onboarded", {
      userId: member.user.id,
      changed: outcome.changed,
      projects: outcome.projects,
      permission: outcome.permission
    });
  } catch (error) {
    log("error", "Guild member auto-onboard failed", {
      userId: member.user.id,
      error: String(error)
    });
  }
});

client.on(Events.InteractionCreate, async (interaction) => {
  if (!interaction.isChatInputCommand()) {
    return;
  }

  const acknowledged = await ensureDeferred(interaction);
  if (!acknowledged) {
    return;
  }

  if (!isAllowedUser(config, interaction.user.id)) {
    log("warn", "Command denied for unassigned Discord user", {
      command: interaction.commandName,
      userId: interaction.user.id
    });
    await replyOnce(interaction, "Access denied for this Discord user.");
    return;
  }

  const scopeKey = createScopeKey(interaction);
  const queuedAt = Date.now();
  const queueBlocked = inflightByScope.has(scopeKey);
  log("info", "Command queued", {
    command: interaction.commandName,
    userId: interaction.user.id,
    scopeKey,
    queueBlocked
  });
  const previous = inflightByScope.get(scopeKey) ?? Promise.resolve();
  const current = previous
    .catch(() => {})
    .then(() => handleInteraction(interaction, scopeKey, queuedAt))
    .catch(async (error) => {
      log("error", "Interaction failed", {
        command: interaction.commandName,
        userId: interaction.user.id,
        scopeKey,
        error: String(error)
      });
      await safelyRespond(async () => {
        await replyOnce(interaction, `Bridge error:\n${String(error)}`);
      }, interaction, "reply failure");
    })
    .finally(() => {
      if (inflightByScope.get(scopeKey) === current) {
        inflightByScope.delete(scopeKey);
      }
    });

  inflightByScope.set(scopeKey, current);
});

await client.login(config.discordBotToken);

async function handleInteraction(interaction, scopeKey, queuedAt) {
  if (adminCommandNames.has(interaction.commandName) && !isAdminUser(config, interaction.user.id)) {
    log("warn", "Admin command denied", {
      command: interaction.commandName,
      userId: interaction.user.id,
      scopeKey
    });
    await replyOnce(interaction, "Access denied. This command is reserved for bridge administrators.");
    return;
  }

  if (
    !adminCommandNames.has(interaction.commandName) &&
    projectRequiredCommandNames.has(interaction.commandName) &&
    !hasAssignedProjects(config, interaction.user.id)
  ) {
    log("warn", "Project command denied without assigned projects", {
      command: interaction.commandName,
      userId: interaction.user.id,
      scopeKey
    });
    await replyOnce(interaction, "No projects are assigned to this Discord user yet.");
    return;
  }

  const scopeState = getAuthorizedScopeState(scopeKey, interaction.user.id);
  return withCommandLogging(interaction, scopeKey, scopeState, queuedAt, async () => {
    switch (interaction.commandName) {
      case "help":
        await replyOnce(interaction, buildHelpMessage());
        return {};
      case "status":
        await replyCodePanel(interaction, "Status", buildStatusMessage(config, scopeState, interaction.user.id));
        return {};
      case "ls": {
        const requestedPath = interaction.options.getString("path", false) ?? "";
        await replyCodePanel(interaction, "ls", listPathLikeLs(config, scopeState, requestedPath));
        return { requestedPath: requestedPath || "." };
      }
      case "resume":
        return handleResume(interaction, scopeKey, scopeState);
      case "projects":
        await replyCodePanel(interaction, "Projects", listTopLevelProjects(config, interaction.user.id));
        return {};
      case "project": {
        const name = interaction.options.getString("name", true);
        const created = createProjectDirectory(config, interaction.user.id, name);
        if (!created.ok) {
          await replyOnce(interaction, created.error);
          return { created: false };
        }
        state.setScope(scopeKey, buildProjectScopeState(config, created.projectName));
        await replyCodePanel(
          interaction,
          "Project Created",
          `Created: ${created.path}\nCurrent project: ${created.projectName}\nSession: reset`
        );
        return { projectName: created.projectName, changedPath: created.path };
      }
      case "project-switch": {
        const name = interaction.options.getString("name", true);
        const switched = switchProjectDirectory(config, interaction.user.id, name);
        if (!switched.ok) {
          await replyOnce(interaction, switched.error);
          return { switched: false };
        }
        state.setScope(scopeKey, buildProjectScopeState(config, switched.projectName));
        await replyCodePanel(
          interaction,
          "Project Switched",
          `Current project: ${switched.projectName}\nCurrent directory: ${switched.path}\nSession: reset`
        );
        return { projectName: switched.projectName, changedPath: switched.path };
      }
      case "mkdir": {
        const requestedPath = interaction.options.getString("path", true);
        const created = createDirectoryWithinProject(config, scopeState, requestedPath);
        if (!created.ok) {
          await replyOnce(interaction, created.error);
          return { requestedPath };
        }
        await replyCodePanel(interaction, "mkdir", `Created directory:\n${created.path}`);
        return { requestedPath, changedPath: created.path };
      }
      case "pwd":
        await replyCodePanel(
          interaction,
          "pwd",
          `Current project: ${scopeState.projectName ?? "none"}\nCurrent directory: ${scopeState.cwd}`
        );
        return {};
      case "cd": {
        const requestedPath = interaction.options.getString("path", true);
        const resolved = resolveScopedProjectPath(config, scopeState, requestedPath);
        if (!resolved.ok) {
          await replyOnce(interaction, resolved.error);
          return { requestedPath };
        }
        state.setScope(scopeKey, { cwd: resolved.path, threadId: null });
        await replyCodePanel(interaction, "cd", `Current directory: ${resolved.path}\nSession: reset`);
        return { requestedPath, changedPath: resolved.path };
      }
      case "reset": {
        const resetState = scopeState.projectName
          ? buildProjectScopeState(config, scopeState.projectName)
          : buildInitialScopeState(config, interaction.user.id);
        state.setScope(scopeKey, resetState);
        await replyCodePanel(interaction, "reset", `Session reset.\nCurrent directory: ${resetState.cwd}`);
        return { reset: true };
      }
      case "admin-user-add":
        return handleAdminUserAdd(interaction);
      case "admin-user-remove":
        return handleAdminUserRemove(interaction);
      case "admin-user-grant-project":
        return handleAdminUserGrantProject(interaction);
      case "admin-user-set-permission":
        return handleAdminUserSetPermission(interaction);
      case "model":
        return handleUserModel(interaction);
      case "reasoning":
        return handleUserReasoning(interaction);
      case "rate-limits":
        return handleRateLimits(interaction);
      case "ask":
        return handleAsk(interaction, scopeKey);
      default:
        await replyOnce(interaction, "Unknown command.");
        return {};
    }
  });
}

async function handleAsk(interaction, scopeKey) {
  const prompt = interaction.options.getString("prompt", true);
  const scopeState = getAuthorizedScopeState(scopeKey, interaction.user.id);
  const attachments = ["attachment1", "attachment2", "attachment3"]
    .map((name) => interaction.options.getAttachment(name, false))
    .filter(Boolean);
  const attachmentContext = await prepareAttachmentsForPrompt(config, scopeState, interaction.user.id, attachments);
  const progressUpdater = createAskProgressUpdater(interaction, scopeState);
  const userSettings = getEffectiveUserSettings(config, interaction.user.id);
  try {
    const result = await runCodex(config, scopeState, prompt, attachmentContext, {
      model: userSettings.model,
      reasoningEffort: userSettings.reasoningEffort,
      sandboxMode: userSettings.sandboxMode,
      onProgress: (progress) => progressUpdater.update(progress)
    });

    if (result.threadId && result.threadId !== scopeState.threadId) {
      state.setScope(scopeKey, { threadId: result.threadId });
    }

    await progressUpdater.finish(result.progress ?? createFallbackAskProgress(scopeState, result));
    return {
      threadId: result.threadId ?? null,
      attachmentCount: attachments.length
    };
  } catch (error) {
    await progressUpdater.finish(createFailedAskProgress(scopeState, error));
    return {
      failed: true,
      attachmentCount: attachments.length
    };
  }
}

async function handleResume(interaction, scopeKey, scopeState) {
  const found = findLatestProjectSession(config, scopeState);
  if (!found.ok) {
    await replyOnce(interaction, found.error);
    return {
      resumed: false
    };
  }

  state.setScope(scopeKey, { threadId: found.session.sessionId });

  const embed = new EmbedBuilder()
    .setColor(0x2b7fff)
    .setTitle(`Resumed ${scopeState.projectName ?? "project"} session`)
    .setDescription(
      [
        `Session: \`${found.session.sessionId}\``,
        found.session.threadName ? `Thread: ${found.session.threadName}` : null,
        found.session.originator ? `Origin: ${found.session.originator}` : null,
        found.session.source ? `Source: ${found.session.source}` : null,
        found.session.updatedAt ? `Updated: ${found.session.updatedAt}` : null
      ]
        .filter(Boolean)
        .join("\n")
    )
    .setFooter({ text: "This Discord scope is now bound to the resumed session. Use /ask to continue." });

  for (const [index, message] of found.session.recentMessages.entries()) {
    embed.addFields({
      name: `${index + 1}. ${formatRoleLabel(message.role, message.phase)}`,
      value: formatResumePreview(message.text, 900)
    });
  }

  if (found.session.recentMessages.length === 0) {
    embed.addFields({
      name: "Recent dialogue",
      value: "No user/assistant messages were found in this session file."
    });
  }

  await replyWithEmbed(interaction, {
    content: `Bound current project to session \`${found.session.sessionId}\`.`,
    embeds: [embed]
  });
  return {
    resumed: true,
    threadId: found.session.sessionId
  };
}

async function withCommandLogging(interaction, scopeKey, scopeState, queuedAt, fn) {
  const startedAt = Date.now();
  const queueDelayMs = Math.max(0, startedAt - queuedAt);
  const baseContext = {
    command: interaction.commandName,
    userId: interaction.user.id,
    scopeKey,
    projectName: scopeState.projectName ?? "none",
    cwd: scopeState.cwd,
    queueDelayMs
  };

  log("info", "Command started", baseContext);

  try {
    const result = (await fn()) ?? {};
    const finalScopeState = getAuthorizedScopeState(scopeKey, interaction.user.id);
    log("info", "Command completed", {
      ...baseContext,
      durationMs: Date.now() - startedAt,
      projectName: finalScopeState.projectName ?? "none",
      cwd: finalScopeState.cwd,
      threadId: finalScopeState.threadId ?? result.threadId ?? null,
      result: result.failed ? "failed" : "ok",
      ...stripEmptyLogFields(result)
    });
    return result;
  } catch (error) {
    log("error", "Command failed", {
      ...baseContext,
      durationMs: Date.now() - startedAt,
      error: String(error)
    });
    throw error;
  }
}

function stripEmptyLogFields(input) {
  const entries = Object.entries(input ?? {}).filter(([, value]) => {
    if (value === null || value === undefined || value === "") {
      return false;
    }
    if (Array.isArray(value) && value.length === 0) {
      return false;
    }
    return true;
  });
  return Object.fromEntries(entries);
}

function getAuthorizedScopeState(scopeKey, userId) {
  const current = state.getScope(scopeKey);
  const reconciled = reconcileScopeState(config, userId, current);
  if (reconciled.changed) {
    state.setScope(scopeKey, reconciled.scopeState);
  }
  return reconciled.scopeState;
}

function formatRoleLabel(role, phase) {
  if (role !== "assistant") {
    return "User";
  }
  return phase ? `Assistant (${phase})` : "Assistant";
}

function createAskProgressUpdater(interaction, scopeState) {
  let lastFingerprint = "";
  let latestProgress = null;
  let sentFullResult = false;
  const timer = setInterval(() => {
    if (!latestProgress || isAskProgressFinished(latestProgress.status)) {
      return;
    }
    void thisUpdater.update(latestProgress);
  }, 1000);

  const thisUpdater = {
    async update(progress) {
      latestProgress = progress;
      const payload = buildAskProgressPayload(scopeState, progress);
      const fingerprint = JSON.stringify({
        title: payload.embeds?.[0]?.data?.title ?? "",
        content: payload.content ?? "",
        status: progress.status,
        threadId: progress.threadId,
        commentary: progress.commentary.map((item) => item.text),
        actions: progress.actions.map((item) => [item.status, item.label, item.text, item.outputPreview ?? ""]),
        finalMessage: progress.finalMessage,
        errorMessage: progress.errorMessage
      });

      if (fingerprint === lastFingerprint) {
        return;
      }

      lastFingerprint = fingerprint;
      await replyWithEmbed(interaction, payload);
    },
    async finish(progress) {
      const effectiveProgress = mergeAskProgress(latestProgress, progress);
      if (!effectiveProgress) {
        return;
      }

       clearInterval(timer);

      await this.update(effectiveProgress);

      const finalBody = effectiveProgress.status === "failed" ? effectiveProgress.errorMessage : effectiveProgress.finalMessage;
      if (!finalBody || sentFullResult) {
        return;
      }

      sentFullResult = true;
      const title = effectiveProgress.status === "failed" ? "Codex Error" : "Codex Result";
      const panels = buildCodePanelMessages(title, finalBody);
      for (const panel of panels) {
        await interaction.followUp(panel);
      }
    }
  };

  return thisUpdater;
}

function mergeAskProgress(baseProgress, overrideProgress) {
  if (!baseProgress) {
    return overrideProgress ?? null;
  }
  if (!overrideProgress) {
    return baseProgress;
  }

  return {
    ...baseProgress,
    ...overrideProgress,
    attachments: overrideProgress.attachments?.length ? overrideProgress.attachments : baseProgress.attachments,
    commentary: overrideProgress.commentary?.length ? overrideProgress.commentary : baseProgress.commentary,
    actions: overrideProgress.actions?.length ? overrideProgress.actions : baseProgress.actions,
    finalMessage: overrideProgress.finalMessage || baseProgress.finalMessage,
    errorMessage: overrideProgress.errorMessage || baseProgress.errorMessage,
    threadId: overrideProgress.threadId || baseProgress.threadId
  };
}

function buildAskProgressPayload(scopeState, progress) {
  const elapsedLabel = formatAskElapsed(progress);
  const statusLabel = getAskStatusLabel(progress.status, elapsedLabel);
  const color = progress.status === "failed" ? 0xc0392b : progress.status === "completed" ? 0x1f8b4c : 0x2b7fff;
  const embed = new EmbedBuilder()
    .setColor(color)
    .setTitle(statusLabel)
    .setDescription(
      [
        `Project: \`${scopeState.projectName ?? "none"}\``,
        `Directory: \`${scopeState.cwd}\``,
        `Session: \`${progress.threadId ?? "pending"}\``,
        progress.attachments.length > 0
          ? `Attachments: ${progress.attachments.map((item) => item.originalName).join(", ")}`
          : null
      ]
        .filter(Boolean)
        .join("\n")
    )
    .setFooter({ text: "Discord bridge mobile view: plan, commands, result." });

  embed.addFields({
    name: "Plan",
    value: formatResumePreview(formatPlanPreview(progress), 900)
  });
  embed.addFields({
    name: "Commands",
    value: formatResumePreview(formatActionPreview(progress.actions), 900)
  });
  embed.addFields({
    name: progress.status === "failed" ? "Error" : "Result",
    value: formatResumePreview(formatResultPreview(progress), 900)
  });

  return {
    content: progress.status === "completed"
      ? `Completed session \`${progress.threadId ?? "unknown"}\` in ${elapsedLabel}.`
      : progress.status === "failed"
        ? `Session \`${progress.threadId ?? "unknown"}\` failed after ${elapsedLabel}.`
        : `Working in \`${scopeState.projectName ?? "project"}\` for ${elapsedLabel}...`,
    embeds: [embed]
  };
}

function createFallbackAskProgress(scopeState, result) {
  const finishedAt = new Date().toISOString();
  return {
    projectName: scopeState.projectName ?? null,
    cwd: scopeState.cwd,
    threadId: result.threadId ?? null,
    status: "completed",
    startedAt: finishedAt,
    finishedAt,
    updatedAt: finishedAt,
    attachments: [],
    commentary: [],
    actions: [],
    finalMessage: result.message ?? "",
    errorMessage: ""
  };
}

function createFailedAskProgress(scopeState, error) {
  const finishedAt = new Date().toISOString();
  return {
    projectName: scopeState.projectName ?? null,
    cwd: scopeState.cwd,
    threadId: scopeState.threadId ?? null,
    status: "failed",
    startedAt: finishedAt,
    finishedAt,
    updatedAt: finishedAt,
    attachments: [],
    commentary: [],
    actions: [],
    finalMessage: "",
    errorMessage: String(error)
  };
}

function getAskStatusLabel(status, elapsedLabel) {
  switch (status) {
    case "completed":
      return `Codex Completed ${elapsedLabel}`;
    case "failed":
      return `Codex Failed ${elapsedLabel}`;
    case "resuming":
    case "running":
    case "starting":
    default:
      return `Codex Thinking ${elapsedLabel}`;
  }
}

function formatAskElapsed(progress) {
  const startedAt = Date.parse(progress.startedAt ?? "");
  const finishedAt = Date.parse(progress.finishedAt ?? "");
  const end = Number.isFinite(finishedAt) ? finishedAt : Date.now();
  const start = Number.isFinite(startedAt) ? startedAt : end;
  return formatElapsedMs(Math.max(0, end - start));
}

function formatElapsedMs(durationMs) {
  const totalSeconds = Math.max(0, Math.floor(durationMs / 1000));
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds % 60;
  if (minutes <= 0) {
    return `${seconds}s`;
  }
  return `${minutes}m ${String(seconds).padStart(2, "0")}s`;
}

function isAskProgressFinished(status) {
  return status === "completed" || status === "failed";
}

function formatPlanPreview(progress) {
  if (progress.commentary.length === 0) {
    return "Waiting for Codex plan...";
  }

  return progress.commentary
    .slice(-3)
    .map((item, index) => `${index + 1}. ${toMobileSingleLine(item.text)}`)
    .join("\n");
}

function formatActionPreview(actions) {
  if (actions.length === 0) {
    return "No commands yet.";
  }

  return actions
    .slice(-5)
    .map((action) => {
      const suffix = action.outputPreview ? ` | ${action.outputPreview}` : "";
      return `${action.status.padEnd(3, " ")} ${action.label}: ${toMobileSingleLine(action.text)}${suffix}`;
    })
    .join("\n");
}

function formatResultPreview(progress) {
  if (progress.status === "failed") {
    return progress.errorMessage || "Codex execution failed.";
  }

  if (!progress.finalMessage) {
    return "Waiting for final result...";
  }

  return progress.finalMessage;
}

function formatResumePreview(text, maxLength) {
  const preview = truncateResumeText(text, maxLength - 11);
  return `\`\`\`text\n${preview}\n\`\`\``;
}

function truncateResumeText(text, maxLength) {
  const normalized = String(text ?? "").trim().replace(/\n{3,}/g, "\n\n").replace(/```/g, "'''");
  if (normalized.length <= maxLength) {
    return normalized;
  }
  return `${normalized.slice(0, maxLength - 3).trimEnd()}...`;
}

function buildCodePanelMessages(title, body, footer = "") {
  const normalizedTitle = String(title ?? "Output").trim() || "Output";
  const chunks = chunkCodeBlockText(body);
  return chunks.map((chunk, index) => {
    const heading = chunks.length === 1 ? normalizedTitle : `${normalizedTitle} (${index + 1}/${chunks.length})`;
    const lines = [`**${heading}**`, `\`\`\`text\n${chunk}\n\`\`\``];
    if (footer && index === chunks.length - 1) {
      lines.push(footer);
    }
    return lines.join("\n");
  });
}

function chunkCodeBlockText(text, maxChars = 1500) {
  const normalized = sanitizeCodeBlockText(text) || "(empty)";
  if (normalized.length <= maxChars) {
    return [normalized];
  }

  const chunks = [];
  let remaining = normalized;

  while (remaining.length > maxChars) {
    let slice = remaining.slice(0, maxChars);
    const splitIndex = slice.lastIndexOf("\n");
    if (splitIndex > maxChars * 0.5) {
      slice = slice.slice(0, splitIndex);
    }
    chunks.push(slice.trimEnd());
    remaining = remaining.slice(slice.length).trimStart();
  }

  if (remaining) {
    chunks.push(remaining);
  }

  return chunks;
}

function sanitizeCodeBlockText(text) {
  return String(text ?? "").replace(/```/g, "'''").trim();
}

function toMobileSingleLine(text, maxLength = 140) {
  const normalized = String(text ?? "").replace(/\s+/g, " ").trim();
  if (normalized.length <= maxLength) {
    return normalized;
  }
  return `${normalized.slice(0, maxLength - 3).trimEnd()}...`;
}

function truncateForEmbed(text, maxLength) {
  const normalized = String(text ?? "").trim().replace(/\n{3,}/g, "\n\n");
  if (normalized.length <= maxLength) {
    return normalized;
  }
  return `${normalized.slice(0, maxLength - 1).trimEnd()}…`;
}

async function handleAdminUserAdd(interaction) {
  const parsed = resolveTargetUserId(interaction);
  if (!parsed.ok) {
    await replyOnce(interaction, parsed.error);
    return { adminAction: "user-add", changed: false };
  }

  const outcome = provisionUserAccess(parsed.userId);
  const resultLabel = outcome.changed ? "Added" : "Already present";
  await replyCodePanel(
    interaction,
    "User Access",
    `User: ${parsed.userId}\nResult: ${resultLabel}\nAssigned projects: ${formatProjectsForMessage(outcome.projects)}\nPermission: ${outcome.permission}`
  );
  return { adminAction: "user-add", changed: outcome.changed, targetUserId: parsed.userId };
}

async function handleAdminUserRemove(interaction) {
  const parsed = resolveTargetUserId(interaction);
  if (!parsed.ok) {
    await replyOnce(interaction, parsed.error);
    return { adminAction: "user-remove", changed: false };
  }

  const nextProjects = cloneUserProjects(config.discordUserProjects);
  const nextPermissions = new Map(config.discordUserPermissions);
  const nextModels = new Map(config.discordUserModels);
  const nextReasoning = new Map(config.discordUserReasoning);
  if (!nextProjects.has(parsed.userId)) {
    await replyOnce(interaction, `User ${parsed.userId} is not present in ACL.`);
    return { adminAction: "user-remove", changed: false, targetUserId: parsed.userId };
  }

  nextProjects.delete(parsed.userId);
  nextPermissions.delete(parsed.userId);
  nextModels.delete(parsed.userId);
  nextReasoning.delete(parsed.userId);
  persistUserAccess(nextProjects, {
    permissions: nextPermissions,
    models: nextModels,
    reasoning: nextReasoning
  });
  state.deleteScopesForUser(parsed.userId);
  await replyOnce(interaction, `Removed user ${parsed.userId} from ACL and cleared any saved bridge scopes and overrides.`);
  return { adminAction: "user-remove", changed: true, targetUserId: parsed.userId };
}

async function handleAdminUserGrantProject(interaction) {
  const parsedUser = resolveTargetUserId(interaction);
  if (!parsedUser.ok) {
    await replyOnce(interaction, parsedUser.error);
    return { adminAction: "grant-project", changed: false };
  }

  const parsedProject = normalizeProjectName(interaction.options.getString("project", true));
  if (!parsedProject.ok) {
    await replyOnce(interaction, parsedProject.error);
    return { adminAction: "grant-project", changed: false, targetUserId: parsedUser.userId };
  }

  const nextProjects = cloneUserProjects(config.discordUserProjects);
  if (!nextProjects.has(parsedUser.userId)) {
    await replyOnce(interaction, `User ${parsedUser.userId} is not present in ACL. Add the user first.`);
    return { adminAction: "grant-project", changed: false, targetUserId: parsedUser.userId, projectName: parsedProject.projectName };
  }

  const currentProjects = nextProjects.get(parsedUser.userId) ?? [];
  if (currentProjects.includes(parsedProject.projectName)) {
    await replyOnce(
      interaction,
      `User ${parsedUser.userId} already has access to ${parsedProject.projectName}.\nAssigned projects: ${formatProjectsForMessage(
        currentProjects
      )}`
    );
    return { adminAction: "grant-project", changed: false, targetUserId: parsedUser.userId, projectName: parsedProject.projectName };
  }

  nextProjects.set(parsedUser.userId, [...currentProjects, parsedProject.projectName]);
  persistUserProjects(nextProjects);
  await replyOnce(
    interaction,
    `Granted ${parsedProject.projectName} to ${parsedUser.userId}.\nAssigned projects: ${formatProjectsForMessage(
      nextProjects.get(parsedUser.userId) ?? []
    )}`
  );
  return { adminAction: "grant-project", changed: true, targetUserId: parsedUser.userId, projectName: parsedProject.projectName };
}

async function handleAdminUserSetPermission(interaction) {
  const parsedUser = resolveTargetUserId(interaction);
  if (!parsedUser.ok) {
    await replyOnce(interaction, parsedUser.error);
    return { adminAction: "set-permission", changed: false };
  }

  if (!config.discordUserProjects.has(parsedUser.userId)) {
    await replyOnce(interaction, `User ${parsedUser.userId} is not present in ACL. Add the user first.`);
    return { adminAction: "set-permission", changed: false, targetUserId: parsedUser.userId };
  }

  const permission = normalizePermissionLevel(interaction.options.getString("permission", true));
  const nextPermissions = new Map(config.discordUserPermissions);
  if (permission === "default") {
    nextPermissions.delete(parsedUser.userId);
  } else {
    nextPermissions.set(parsedUser.userId, permission);
  }

  updateEnvVariable("DISCORD_USER_PERMISSIONS", serializeKeyValueMap(nextPermissions));
  reloadRuntimeConfig();
  await replyCodePanel(
    interaction,
    "Permission Updated",
    `User: ${parsedUser.userId}\nPermission: ${permission}\nExecution mode: ${getEffectiveUserSettings(config, parsedUser.userId).sandboxMode}`
  );
  return { adminAction: "set-permission", changed: true, targetUserId: parsedUser.userId, permission };
}

async function handleUserModel(interaction) {
  const requested = (interaction.options.getString("value", false) ?? "").trim();
  if (!requested) {
    const current = getEffectiveUserSettings(config, interaction.user.id).model || "default";
    await replyCodePanel(interaction, "Model", `Current model: ${current}`);
    return { setting: "model", value: current };
  }

  const nextModels = new Map(config.discordUserModels);
  if (requested.toLowerCase() === "default") {
    nextModels.delete(interaction.user.id);
  } else {
    nextModels.set(interaction.user.id, requested);
  }
  updateEnvVariable("DISCORD_USER_MODELS", serializeKeyValueMap(nextModels));
  reloadRuntimeConfig();
  const current = getEffectiveUserSettings(config, interaction.user.id).model || "default";
  await replyCodePanel(interaction, "Model Updated", `Current model: ${current}`);
  return { setting: "model", value: current };
}

async function handleUserReasoning(interaction) {
  const requested = normalizeReasoningValue(interaction.options.getString("value", false) ?? "");
  if (!requested) {
    const current = getEffectiveUserSettings(config, interaction.user.id).reasoningEffort || "default";
    await replyCodePanel(interaction, "Reasoning", `Current reasoning: ${current}`);
    return { setting: "reasoning", value: current };
  }

  const nextReasoning = new Map(config.discordUserReasoning);
  if (requested === "default") {
    nextReasoning.delete(interaction.user.id);
  } else {
    nextReasoning.set(interaction.user.id, requested);
  }
  updateEnvVariable("DISCORD_USER_REASONING", serializeKeyValueMap(nextReasoning));
  reloadRuntimeConfig();
  const current = getEffectiveUserSettings(config, interaction.user.id).reasoningEffort || "default";
  await replyCodePanel(interaction, "Reasoning Updated", `Current reasoning: ${current}`);
  return { setting: "reasoning", value: current };
}

async function handleRateLimits(interaction) {
  const latest = readLatestRateLimits(config);
  if (!latest.ok) {
    await replyCodePanel(interaction, "Rate Limits", latest.error);
    return { setting: "rate-limits", found: false };
  }

  await replyCodePanel(interaction, "Rate Limits", formatRateLimitsMessage(latest));
  return { setting: "rate-limits", found: true };
}

function persistUserProjects(nextProjects) {
  persistUserAccess(nextProjects);
}

function persistUserAccess(nextProjects, overrides = {}) {
  updateEnvVariable("DISCORD_USER_PROJECTS", serializeUserProjects(nextProjects));
  if (overrides.permissions) {
    updateEnvVariable("DISCORD_USER_PERMISSIONS", serializeKeyValueMap(overrides.permissions));
  }
  if (overrides.models) {
    updateEnvVariable("DISCORD_USER_MODELS", serializeKeyValueMap(overrides.models));
  }
  if (overrides.reasoning) {
    updateEnvVariable("DISCORD_USER_REASONING", serializeKeyValueMap(overrides.reasoning));
  }
  reloadRuntimeConfig();
}

function autoOnboardGuildMember(userId) {
  return provisionUserAccess(userId);
}

function provisionUserAccess(userId) {
  const nextProjects = cloneUserProjects(config.discordUserProjects);
  const nextPermissions = new Map(config.discordUserPermissions);
  const currentProjects = nextProjects.get(userId) ?? [];
  const mergedProjects = [...new Set([...currentProjects, ...config.discordAutoAddProjects])];
  const alreadyPresent = nextProjects.has(userId);
  const projectsChanged = mergedProjects.length !== currentProjects.length;
  const nextPermission = normalizePermissionLevel(config.discordAutoAddPermission);
  const currentPermission = normalizePermissionLevel(config.discordUserPermissions.get(userId) ?? "default");
  const permissionChanged = nextPermission === "full-access" && currentPermission !== "full-access";

  if (!alreadyPresent || projectsChanged) {
    nextProjects.set(userId, mergedProjects);
  }
  if (permissionChanged) {
    nextPermissions.set(userId, nextPermission);
  }

  const changed = !alreadyPresent || projectsChanged || permissionChanged;
  if (changed) {
    persistUserAccess(nextProjects, permissionChanged ? { permissions: nextPermissions } : {});
  }

  return {
    changed,
    projects: mergedProjects,
    permission: permissionChanged ? nextPermission : currentPermission
  };
}

function resolveTargetUserId(interaction) {
  const user = interaction.options.getUser("user", false);
  if (user?.id) {
    return { ok: true, userId: user.id };
  }

  const rawUserId = interaction.options.getString("user_id", false)?.trim() ?? "";
  if (rawUserId) {
    return normalizeDiscordUserId(rawUserId);
  }

  return { ok: false, error: "Provide a Discord user mention or a numeric user_id." };
}

function reloadRuntimeConfig() {
  dotenv.config({ path: envFilePath, override: true });
  config = loadConfig();
  validateRuntimeConfig(config);
}

function updateEnvVariable(key, value) {
  const fileText = fs.existsSync(envFilePath) ? fs.readFileSync(envFilePath, "utf8") : "";
  const lines = fileText === "" ? [] : fileText.split(/\r?\n/);
  const nextLine = `${key}=${value}`;
  let replaced = false;

  const updated = lines.map((line) => {
    if (line.startsWith(`${key}=`)) {
      replaced = true;
      return nextLine;
    }
    return line;
  });

  if (!replaced) {
    updated.push(nextLine);
  }

  const normalized = updated.join("\n").replace(/\n+$/u, "");
  fs.writeFileSync(envFilePath, `${normalized}\n`, "utf8");
}

function formatProjectsForMessage(projects) {
  return projects.length === 0 ? "none" : projects.join(", ");
}

function formatRateLimitsMessage(latest) {
  const primary = latest.rateLimits?.primary ?? {};
  const secondary = latest.rateLimits?.secondary ?? {};
  return [
    `Observed at: ${formatRateLimitTime(latest.timestamp)}`,
    "",
    `Primary remaining: ${formatRemainingPercent(primary.used_percent)}`,
    `Primary resets: ${formatRateLimitTime(primary.resets_at)}`,
    "",
    `Secondary remaining: ${formatRemainingPercent(secondary.used_percent)}`,
    `Secondary resets: ${formatRateLimitTime(secondary.resets_at)}`
  ].join("\n");
}

function formatRemainingPercent(usedPercent) {
  const used = Number.isFinite(usedPercent) ? usedPercent : Number.parseFloat(usedPercent ?? "0");
  if (!Number.isFinite(used)) {
    return "unknown";
  }
  const remaining = Math.max(0, 100 - used);
  return `${remaining.toFixed(1)}%`;
}

function formatRateLimitTime(input) {
  const value = Number.isFinite(input) ? Number(input) : Date.parse(String(input ?? ""));
  if (!Number.isFinite(value) || value <= 0) {
    return "unknown";
  }

  const ms = value > 1e12 ? value : value * 1000;
  return new Date(ms).toLocaleString("en-AU", {
    timeZone: "Australia/Sydney",
    hour12: false
  });
}

async function ensureDeferred(interaction) {
  if (interaction.deferred || interaction.replied) {
    return true;
  }

  try {
    await interaction.deferReply();
    return true;
  } catch (error) {
    if (isIgnorableDiscordResponseError(error)) {
      log("warn", "Dropping expired interaction before handling", {
        command: interaction.commandName,
        userId: interaction.user.id,
        error: String(error)
      });
      return false;
    }
    throw error;
  }
}

async function replyOnce(interaction, content) {
  await replyParts(interaction, chunkDiscordMessage(content));
}

async function replyCodePanel(interaction, title, body, footer = "") {
  await replyParts(interaction, buildCodePanelMessages(title, body, footer));
}

async function replyWithEmbed(interaction, payload) {
  await safelyRespond(async () => {
    if (interaction.deferred || interaction.replied) {
      await interaction.editReply(payload);
      return;
    }

    await interaction.reply(payload);
  }, interaction, "embed reply");
}

async function replyParts(interaction, parts) {
  if (parts.length === 0) {
    parts = ["(empty response)"];
  }

  await safelyRespond(async () => {
    if (interaction.deferred || interaction.replied) {
      await interaction.editReply(parts[0]);
      for (const part of parts.slice(1)) {
        await interaction.followUp(part);
      }
      return;
    }

    await interaction.reply(parts[0]);
    for (const part of parts.slice(1)) {
      await interaction.followUp(part);
    }
  }, interaction, "text reply");
}

async function safelyRespond(fn, interaction, label) {
  try {
    await fn();
  } catch (error) {
    if (isIgnorableDiscordResponseError(error)) {
      log("warn", "Ignoring expired Discord interaction response", {
        label,
        command: interaction.commandName,
        userId: interaction.user.id,
        error: String(error)
      });
      return;
    }
    throw error;
  }
}

function isIgnorableDiscordResponseError(error) {
  return error?.code === 10062 || error?.code === 40060;
}
