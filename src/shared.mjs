import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawn } from "node:child_process";

import { REST, Routes, SlashCommandBuilder } from "discord.js";

export function loadConfig(env = process.env) {
  const config = {
    discordBotToken: env.DISCORD_BOT_TOKEN ?? "",
    discordApplicationId: env.DISCORD_APPLICATION_ID ?? "",
    discordGuildId: env.DISCORD_GUILD_ID ?? "",
    discordAdminUserIds: parseCsv(env.DISCORD_ADMIN_USER_IDS ?? ""),
    discordAllowedUserIds: parseCsv(env.DISCORD_ALLOWED_USER_IDS ?? ""),
    discordUserProjects: parseUserProjects(env.DISCORD_USER_PROJECTS ?? ""),
    discordUserPermissions: parseKeyValueMap(env.DISCORD_USER_PERMISSIONS ?? ""),
    discordUserModels: parseKeyValueMap(env.DISCORD_USER_MODELS ?? ""),
    discordUserReasoning: parseKeyValueMap(env.DISCORD_USER_REASONING ?? ""),
    discordAutoAddOnGuildJoin: parseBoolean(env.DISCORD_AUTO_ADD_ON_GUILD_JOIN ?? "false"),
    discordAutoAddProjects: parseProjectList(env.DISCORD_AUTO_ADD_PROJECTS ?? ""),
    discordAutoAddPermission: normalizePermissionLevel(env.DISCORD_AUTO_ADD_PERMISSION ?? "default"),
    workspaceRoot: normalizeAbsolutePosix(env.WORKSPACE_ROOT ?? "/mnt/d/codex_workspace"),
    defaultCwd: normalizeAbsolutePosix(env.DEFAULT_CWD ?? env.WORKSPACE_ROOT ?? "/mnt/d/codex_workspace"),
    stateDir: resolveStateDir(env.STATE_DIR ?? "./data"),
    codexHome: (env.CODEX_HOME ?? "").trim(),
    codexModel: env.CODEX_MODEL?.trim() || "",
    codexExecMode: (env.CODEX_EXEC_MODE ?? "read-only").trim(),
    codexTimeoutMs: Number.parseInt(env.CODEX_TIMEOUT_MS ?? "900000", 10),
    codexMessagePrefix: (env.CODEX_MESSAGE_PREFIX ?? "").trim(),
    discordMaxAttachmentBytes: Number.parseInt(env.DISCORD_MAX_ATTACHMENT_BYTES ?? "20971520", 10)
  };

  if (!path.posix.isAbsolute(config.workspaceRoot) || !path.posix.isAbsolute(config.defaultCwd)) {
    throw new Error("WORKSPACE_ROOT and DEFAULT_CWD must be absolute Linux paths visible inside WSL.");
  }

  validateConfiguredProjects(config.discordUserProjects);
  validateProjectNames(config.discordAutoAddProjects, "DISCORD_AUTO_ADD_PROJECTS");
  return config;
}

export function validateRuntimeConfig(config) {
  if (!config.discordBotToken) {
    throw new Error("Missing DISCORD_BOT_TOKEN.");
  }
}

export function validateRegisterConfig(config) {
  if (!config.discordBotToken || !config.discordApplicationId || !config.discordGuildId) {
    throw new Error("Missing DISCORD_BOT_TOKEN, DISCORD_APPLICATION_ID, or DISCORD_GUILD_ID.");
  }
}

export function buildSlashCommands() {
  return [
    new SlashCommandBuilder().setName("help").setDescription("Show available bridge commands"),
    new SlashCommandBuilder().setName("status").setDescription("Show current project, directory, and session status"),
    new SlashCommandBuilder()
      .setName("ls")
      .setDescription("List files like ls -la inside the current project")
      .addStringOption((option) =>
        option.setName("path").setDescription("Optional path inside the current project").setRequired(false)
      ),
    new SlashCommandBuilder()
      .setName("ask")
      .setDescription("Send a prompt to Codex inside the current project")
      .addStringOption((option) =>
        option.setName("prompt").setDescription("What Codex should do").setRequired(true)
      )
      .addAttachmentOption((option) =>
        option.setName("attachment1").setDescription("Optional file to include with the prompt").setRequired(false)
      )
      .addAttachmentOption((option) =>
        option.setName("attachment2").setDescription("Optional second file to include with the prompt").setRequired(false)
      )
      .addAttachmentOption((option) =>
        option.setName("attachment3").setDescription("Optional third file to include with the prompt").setRequired(false)
      ),
    new SlashCommandBuilder().setName("resume").setDescription("Resume the most recent Codex session for the current project"),
    new SlashCommandBuilder().setName("projects").setDescription("List projects assigned to this Discord user"),
    new SlashCommandBuilder()
      .setName("project")
      .setDescription("Create an assigned top-level project directory and switch into it")
      .addStringOption((option) =>
        option.setName("name").setDescription("Assigned top-level project directory name").setRequired(true)
      ),
    new SlashCommandBuilder()
      .setName("project-switch")
      .setDescription("Switch to an assigned top-level project and reset the current session")
      .addStringOption((option) =>
        option.setName("name").setDescription("Assigned top-level project directory name").setRequired(true)
      ),
    new SlashCommandBuilder()
      .setName("mkdir")
      .setDescription("Create a directory inside the current project")
      .addStringOption((option) =>
        option.setName("path").setDescription("Relative or absolute path inside the current project").setRequired(true)
      ),
    new SlashCommandBuilder()
      .setName("cd")
      .setDescription("Change the working directory inside the current project")
      .addStringOption((option) =>
        option.setName("path").setDescription("Relative or absolute path inside the current project").setRequired(true)
      ),
    new SlashCommandBuilder().setName("pwd").setDescription("Show the current project directory"),
    new SlashCommandBuilder().setName("reset").setDescription("Reset the Codex session and return to the project root"),
    new SlashCommandBuilder()
      .setName("admin-user-add")
      .setDescription("Admin: add a Discord user to bridge access")
      .addUserOption((option) =>
        option.setName("user").setDescription("Discord user mention").setRequired(false)
      )
      .addStringOption((option) =>
        option.setName("user_id").setDescription("Discord user id fallback").setRequired(false)
      ),
    new SlashCommandBuilder()
      .setName("admin-user-remove")
      .setDescription("Admin: remove a Discord user from bridge access")
      .addUserOption((option) =>
        option.setName("user").setDescription("Discord user mention").setRequired(false)
      )
      .addStringOption((option) =>
        option.setName("user_id").setDescription("Discord user id fallback").setRequired(false)
      ),
    new SlashCommandBuilder()
      .setName("admin-user-grant-project")
      .setDescription("Admin: grant one project to a Discord user")
      .addStringOption((option) =>
        option.setName("project").setDescription("Top-level project name").setRequired(true)
      )
      .addUserOption((option) =>
        option.setName("user").setDescription("Discord user mention").setRequired(false)
      )
      .addStringOption((option) =>
        option.setName("user_id").setDescription("Discord user id fallback").setRequired(false)
      ),
    new SlashCommandBuilder()
      .setName("admin-user-set-permission")
      .setDescription("Admin: set default or full-access for a Discord user")
      .addStringOption((option) =>
        option
          .setName("permission")
          .setDescription("Permission level")
          .setRequired(true)
          .addChoices(
            { name: "default", value: "default" },
            { name: "full-access", value: "full-access" }
          )
      )
      .addUserOption((option) =>
        option.setName("user").setDescription("Discord user mention").setRequired(false)
      )
      .addStringOption((option) =>
        option.setName("user_id").setDescription("Discord user id fallback").setRequired(false)
      ),
    new SlashCommandBuilder()
      .setName("model")
      .setDescription("Show or set your Codex model")
      .addStringOption((option) =>
        option.setName("value").setDescription("Example: gpt-5.4 or gpt-5.4-mini").setRequired(false)
      ),
    new SlashCommandBuilder()
      .setName("reasoning")
      .setDescription("Show or set your reasoning effort")
      .addStringOption((option) =>
        option
          .setName("value")
          .setDescription("Reasoning level")
          .setRequired(false)
          .addChoices(
            { name: "default", value: "default" },
            { name: "low", value: "low" },
            { name: "medium", value: "medium" },
            { name: "high", value: "high" }
          )
      ),
    new SlashCommandBuilder().setName("rate-limits").setDescription("Show latest Codex rate-limit remaining estimate")
  ];
}

export async function registerSlashCommands(config) {
  validateRegisterConfig(config);
  const rest = new REST({ version: "10" }).setToken(config.discordBotToken);
  const commands = buildSlashCommands().map((command) => command.toJSON());

  await rest.put(
    Routes.applicationGuildCommands(config.discordApplicationId, config.discordGuildId),
    { body: commands }
  );

  return commands.length;
}

export function createStateStore(stateDir, defaultCwd) {
  fs.mkdirSync(stateDir, { recursive: true });
  const filePath = path.join(stateDir, "bridge-state.json");
  const initial = loadJson(filePath, { scopes: {} });

  function persist() {
    fs.writeFileSync(filePath, JSON.stringify({ scopes: store.scopes }, null, 2), "utf8");
  }

  const store = {
    filePath,
    scopes: initial.scopes ?? {},
    getScope(scopeKey) {
      if (!store.scopes[scopeKey]) {
        store.scopes[scopeKey] = { cwd: defaultCwd, threadId: null, projectName: null };
        persist();
      }
      return store.scopes[scopeKey];
    },
    setScope(scopeKey, next) {
      store.scopes[scopeKey] = { ...store.getScope(scopeKey), ...next };
      persist();
    },
    resetScope(scopeKey) {
      store.scopes[scopeKey] = { cwd: defaultCwd, threadId: null, projectName: null };
      persist();
    },
    deleteScopesForUser(userId) {
      const suffix = `:user:${userId}`;
      let changed = false;
      for (const scopeKey of Object.keys(store.scopes)) {
        if (scopeKey === `dm:${userId}` || scopeKey.endsWith(suffix)) {
          delete store.scopes[scopeKey];
          changed = true;
        }
      }
      if (changed) {
        persist();
      }
    }
  };

  return store;
}

export function buildInitialScopeState(config, userId) {
  const defaultProjectName = selectDefaultProjectName(config, userId);
  if (defaultProjectName) {
    return buildProjectScopeState(config, defaultProjectName);
  }

  return {
    projectName: inferProjectNameFromPath(config, userId, config.defaultCwd),
    cwd: config.defaultCwd,
    threadId: null
  };
}

export function buildProjectScopeState(config, projectName) {
  return {
    projectName,
    cwd: getProjectRoot(config, projectName),
    threadId: null
  };
}

export function reconcileScopeState(config, userId, scopeState) {
  const initial = buildInitialScopeState(config, userId);
  const previous = {
    projectName: typeof scopeState?.projectName === "string" ? scopeState.projectName : null,
    cwd: normalizeScopeCwd(scopeState?.cwd, initial.cwd),
    threadId: scopeState?.threadId ?? null
  };

  if (!hasProjectAcl(config)) {
    const safeCwd = isWithinRoot(previous.cwd, config.workspaceRoot) ? previous.cwd : initial.cwd;
    const safeProjectName = inferProjectNameFromPath(config, userId, safeCwd);
    const next = {
      projectName: safeProjectName,
      cwd: safeCwd,
      threadId: safeCwd === previous.cwd ? previous.threadId : null
    };
    return { changed: !sameScopeState(previous, next), scopeState: next };
  }

  const allowedProjects = getAllowedProjectsForUser(config, userId);
  let projectName = previous.projectName;
  if (!projectName || !allowedProjects.includes(projectName)) {
    projectName = inferProjectNameFromPath(config, userId, previous.cwd) ?? initial.projectName;
  }

  if (!projectName) {
    const next = { ...initial };
    return { changed: !sameScopeState(previous, next), scopeState: next };
  }

  const projectRoot = getProjectRoot(config, projectName);
  const safeCwd = isWithinRoot(previous.cwd, projectRoot) ? previous.cwd : projectRoot;
  const next = {
    projectName,
    cwd: safeCwd,
    threadId: safeCwd === previous.cwd && previous.projectName === projectName ? previous.threadId : null
  };

  return { changed: !sameScopeState(previous, next), scopeState: next };
}

export function resolveScopedProjectPath(config, scopeState, input, options = {}) {
  return resolvePathWithinRoot(
    getScopeProjectRoot(config, scopeState),
    scopeState.cwd,
    input,
    {
      mustExist: options.mustExist ?? true,
      allowFiles: options.allowFiles ?? false
    }
  );
}

export function createDirectoryWithinProject(config, scopeState, input) {
  const candidate = resolveScopedProjectPath(config, scopeState, input, {
    mustExist: false,
    allowFiles: false
  });
  if (!candidate.ok) {
    return candidate;
  }

  if (fs.existsSync(candidate.path)) {
    const stat = fs.statSync(candidate.path);
    if (stat.isDirectory()) {
      return { ok: false, error: `Directory already exists: ${candidate.path}` };
    }
    return { ok: false, error: `Path exists and is not a directory: ${candidate.path}` };
  }

  fs.mkdirSync(candidate.path, { recursive: true });
  return { ok: true, path: candidate.path };
}

export function createProjectDirectory(config, userId, name) {
  const validation = validateProjectName(name);
  if (!validation.ok) {
    return validation;
  }
  if (!isProjectAllowedForUser(config, userId, validation.name)) {
    return { ok: false, error: `Project ${validation.name} is not assigned to this Discord user.` };
  }

  const candidate = getProjectRoot(config, validation.name);
  if (fs.existsSync(candidate)) {
    return { ok: false, error: `Project already exists: ${candidate}` };
  }

  fs.mkdirSync(candidate, { recursive: true });
  return { ok: true, path: candidate, projectName: validation.name };
}

export function switchProjectDirectory(config, userId, name) {
  const validation = validateProjectName(name);
  if (!validation.ok) {
    return validation;
  }
  if (!isProjectAllowedForUser(config, userId, validation.name)) {
    return { ok: false, error: `Project ${validation.name} is not assigned to this Discord user.` };
  }

  const candidate = getProjectRoot(config, validation.name);
  try {
    const stat = fs.statSync(candidate);
    if (!stat.isDirectory()) {
      return { ok: false, error: `${candidate} is not a directory.` };
    }
  } catch {
    return { ok: false, error: `Project does not exist: ${candidate}` };
  }

  return { ok: true, path: candidate, projectName: validation.name };
}

export function listTopLevelProjects(config, userId) {
  const projectNames = hasProjectAcl(config)
    ? getAllowedProjectsForUser(config, userId)
    : listWorkspaceProjects(config);

  if (projectNames.length === 0) {
    return hasProjectAcl(config)
      ? "No projects are assigned to this Discord user."
      : `No projects found under ${config.workspaceRoot}`;
  }

  const lines = projectNames.map((projectName) => {
    const projectRoot = getProjectRoot(config, projectName);
    const exists = isExistingDirectory(projectRoot);
    return exists ? `- ${projectName}` : `- ${projectName} (not created yet)`;
  });

  return ["Projects:", ...lines].join("\n");
}

export function buildStatusMessage(config, scopeState, userId) {
  const projectRoot = getScopeProjectRoot(config, scopeState);
  const allowedProjects = getAllowedProjectsForUser(config, userId);
  const settings = getEffectiveUserSettings(config, userId);

  return [
    `Current project: ${scopeState.projectName ?? "none"}`,
    `Project root: ${projectRoot}`,
    `Current directory: ${scopeState.cwd}`,
    `Permission: ${settings.permission}`,
    `Execution mode: ${settings.sandboxMode}`,
    `Model: ${settings.model || "default"}`,
    `Reasoning: ${settings.reasoningEffort || "default"}`,
    `Active session: ${scopeState.threadId ?? "none"}`,
    allowedProjects.length > 0 ? `Assigned projects: ${allowedProjects.join(", ")}` : null
  ]
    .filter(Boolean)
    .join("\n");
}

export function buildHelpMessage() {
  return [
    "Commands:",
    "/help - show command help",
    "/status - show current project, directory, mode, and session state",
    "/ls [path] - list files like ls -la inside the current project; defaults to current directory",
    "/ask <prompt> [attachment1..3] - send a prompt to Codex with writes limited to the current project root",
    "/resume - bind this Discord scope to the latest Codex session for the current project and preview recent dialogue",
    "/projects - list projects assigned to your Discord user",
    "/project <name> - create an assigned top-level project and switch into it",
    "/project-switch <name> - switch to an assigned existing top-level project and reset the session",
    "/mkdir <path> - create a directory inside the current project",
    "/cd <path> - change the current working directory inside the current project",
    "/pwd - show the current project directory",
    "/reset - clear the current Codex session and return to the project root",
    "/admin-user-add <@user|user_id> - admin only, add a Discord user to ACL",
    "/admin-user-remove <@user|user_id> - admin only, remove a Discord user from ACL",
    "/admin-user-grant-project <@user|user_id> <project> - admin only, grant one project to a user",
    "/admin-user-set-permission <@user|user_id> <default|full-access> - admin only, change execution permission",
    "/model [value] - show or set your model override",
    "/reasoning [low|medium|high] - show or set your reasoning effort override",
    "/rate-limits - show latest remaining rate-limit estimate"
  ].join("\n");
}

export async function runCodex(
  config,
  scopeState,
  promptText,
  attachmentContext = createEmptyAttachmentContext(),
  options = {}
) {
  const projectRoot = getScopeProjectRoot(config, scopeState);
  const cwdReady = fs.existsSync(scopeState.cwd) && fs.statSync(scopeState.cwd).isDirectory();
  const projectReady = fs.existsSync(projectRoot) && fs.statSync(projectRoot).isDirectory();

  if (!projectReady) {
    throw new Error(`Current project root does not exist: ${projectRoot}`);
  }
  if (!cwdReady) {
    throw new Error(`Current directory does not exist: ${scopeState.cwd}`);
  }

  const tmpFile = path.join(os.tmpdir(), `codex-discord-${Date.now()}-${Math.random().toString(16).slice(2)}.txt`);
  const prompt = buildCodexPrompt(config, scopeState, promptText, projectRoot, attachmentContext);
  const args = buildCodexArgs(
    config,
    scopeState,
    prompt,
    outputFilePath(tmpFile, projectRoot),
    attachmentContext,
    options
  );

  return new Promise((resolve, reject) => {
    const progress = createCodexProgress(scopeState, attachmentContext);
    const child = spawn("codex", args, {
      cwd: projectRoot,
      env: buildCodexEnv(config),
      stdio: ["ignore", "pipe", "pipe"]
    });

    let stdout = "";
    let stdoutBuffer = "";
    let stderr = "";
    let threadId = scopeState.threadId;
    let progressDirty = true;
    let progressQueue = Promise.resolve();
    let sessionTail = createSessionTailState(config, threadId, Boolean(scopeState.threadId));

    const timer = setTimeout(() => child.kill("SIGTERM"), config.codexTimeoutMs);
    const progressInterval = options.onProgress
      ? setInterval(() => {
          void flushProgress();
        }, 1500)
      : null;

    if (options.onProgress) {
      void flushProgress(true);
    }

    child.stdout.on("data", (chunk) => {
      const textChunk = chunk.toString("utf8");
      stdout += textChunk;
      stdoutBuffer += textChunk;
      processStdoutBuffer();
    });

    child.stderr.on("data", (chunk) => {
      stderr += chunk.toString("utf8");
    });

    child.on("close", async (code, signal) => {
      clearTimeout(timer);
      if (progressInterval) {
        clearInterval(progressInterval);
      }

      pollSessionTail();

      let finalMessage = "";
      try {
        if (fs.existsSync(tmpFile)) {
          finalMessage = fs.readFileSync(tmpFile, "utf8").trim();
        }
      } finally {
        try {
          fs.unlinkSync(tmpFile);
        } catch {
          // Ignore cleanup failures.
        }
      }

      progress.threadId = threadId;
      progress.status = code === 0 ? "completed" : "failed";
      progress.finishedAt = new Date().toISOString();
      progress.finalMessage = finalMessage || progress.finalMessage;
      progress.errorMessage = code === 0 ? "" : buildCodexFailureDetails(stdout, stderr, threadId, code, signal, finalMessage);
      markProgressChanged();
      await flushProgress(true);

      if (code === 0 && finalMessage) {
        resolve({ ok: true, threadId, message: finalMessage, progress: cloneCodexProgress(progress) });
        return;
      }

      reject(new Error(progress.errorMessage));
    });

    function processStdoutBuffer() {
      const lines = stdoutBuffer.split(/\r?\n/);
      stdoutBuffer = lines.pop() ?? "";

      for (const rawLine of lines) {
        const line = rawLine.trim();
        if (!line.startsWith("{")) {
          continue;
        }

        try {
          const event = JSON.parse(line);
          handleStdoutEvent(event);
        } catch {
          // Ignore non-JSON lines.
        }
      }
    }

    function handleStdoutEvent(event) {
      if (event.type === "thread.started" && event.thread_id) {
        threadId = event.thread_id;
        progress.threadId = threadId;
        progress.status = "running";
        sessionTail = createSessionTailState(config, threadId, false);
        markProgressChanged();
        pollSessionTail();
        return;
      }

      if (event.type === "turn.started") {
        progress.status = "running";
        markProgressChanged();
        return;
      }

      if (event.type === "item.completed" && event.item?.type === "agent_message" && event.item.text) {
        progress.finalMessage = String(event.item.text).trim() || progress.finalMessage;
        markProgressChanged();
      }
    }

    function pollSessionTail() {
      if (!threadId) {
        return;
      }

      if (!sessionTail.filePath) {
        attachSessionTail(threadId, sessionTail.resumeExisting);
      }
      if (!sessionTail.filePath) {
        return;
      }

      let stat;
      try {
        stat = fs.statSync(sessionTail.filePath);
      } catch {
        return;
      }

      if (stat.size <= sessionTail.offset) {
        return;
      }

      const handle = fs.openSync(sessionTail.filePath, "r");
      try {
        const size = stat.size - sessionTail.offset;
        const buffer = Buffer.alloc(size);
        const bytesRead = fs.readSync(handle, buffer, 0, size, sessionTail.offset);
        sessionTail.offset += bytesRead;
        const chunkText = sessionTail.partial + buffer.toString("utf8", 0, bytesRead);
        const lines = chunkText.split(/\r?\n/);
        sessionTail.partial = lines.pop() ?? "";
        for (const rawLine of lines) {
          updateProgressFromSessionLine(progress, rawLine, sessionTail.actionIndexByCallId);
        }
        markProgressChanged();
      } finally {
        fs.closeSync(handle);
      }
    }

    function attachSessionTail(activeThreadId, resumeExisting) {
      const filePath = findSessionFileById(config, activeThreadId);
      if (!filePath) {
        return false;
      }

      sessionTail = {
        threadId: activeThreadId,
        filePath,
        offset: resumeExisting ? fs.statSync(filePath).size : 0,
        partial: "",
        resumeExisting,
        actionIndexByCallId: new Map()
      };
      return true;
    }

    function markProgressChanged() {
      progress.updatedAt = new Date().toISOString();
      progressDirty = true;
    }

    function flushProgress(force = false) {
      pollSessionTail();
      if (!options.onProgress || (!force && !progressDirty)) {
        return progressQueue;
      }

      progressDirty = false;
      const snapshot = cloneCodexProgress(progress);
      progressQueue = progressQueue
        .catch(() => {})
        .then(() => options.onProgress(snapshot));
      return progressQueue;
    }
  });
}

function createCodexProgress(scopeState, attachmentContext) {
  return {
    projectName: scopeState.projectName ?? null,
    cwd: scopeState.cwd,
    threadId: scopeState.threadId ?? null,
    status: scopeState.threadId ? "resuming" : "starting",
    startedAt: new Date().toISOString(),
    finishedAt: null,
    updatedAt: new Date().toISOString(),
    attachments: (attachmentContext.savedFiles ?? []).map((file) => ({
      path: file.path,
      kind: file.kind,
      originalName: file.originalName
    })),
    commentary: [],
    actions: [],
    finalMessage: "",
    errorMessage: ""
  };
}

function cloneCodexProgress(progress) {
  return {
    ...progress,
    attachments: progress.attachments.map((file) => ({ ...file })),
    commentary: progress.commentary.map((item) => ({ ...item })),
    actions: progress.actions.map((item) => ({ ...item }))
  };
}

function createSessionTailState(config, threadId, resumeExisting) {
  const filePath = threadId ? findSessionFileById(config, threadId) : null;
  return {
    threadId,
    filePath,
    offset: filePath && resumeExisting ? fs.statSync(filePath).size : 0,
    partial: "",
    resumeExisting,
    actionIndexByCallId: new Map()
  };
}

function updateProgressFromSessionLine(progress, rawLine, actionIndexByCallId) {
  const trimmed = String(rawLine ?? "").trim();
  if (!trimmed.startsWith("{")) {
    return;
  }

  let parsed;
  try {
    parsed = JSON.parse(trimmed);
  } catch {
    return;
  }

  if (parsed.type !== "response_item") {
    return;
  }

  const payload = parsed.payload ?? {};
  switch (payload.type) {
    case "message":
      updateProgressFromMessage(progress, payload);
      return;
    case "function_call":
      updateProgressFromFunctionCall(progress, payload, actionIndexByCallId);
      return;
    case "function_call_output":
      updateProgressFromFunctionOutput(progress, payload, actionIndexByCallId);
      return;
    default:
      return;
  }
}

function updateProgressFromMessage(progress, payload) {
  if (payload.role !== "assistant") {
    return;
  }

  const text = extractMessageText(payload.content ?? []);
  if (!text || text.startsWith("<environment_context>")) {
    return;
  }

  if (payload.phase === "commentary") {
    const last = progress.commentary.at(-1);
    if (last?.text !== text) {
      progress.commentary.push({
        phase: payload.phase ?? null,
        text
      });
      progress.commentary = progress.commentary.slice(-4);
    }
    return;
  }

  if (payload.phase === "final_answer") {
    progress.finalMessage = text;
  }
}

function updateProgressFromFunctionCall(progress, payload, actionIndexByCallId) {
  const action = summarizeFunctionCall(payload);
  const nextIndex = progress.actions.push(action) - 1;
  if (action.callId) {
    actionIndexByCallId.set(action.callId, nextIndex);
  }
  progress.actions = progress.actions.slice(-8);

  if (progress.actions.length > 0) {
    actionIndexByCallId.clear();
    for (const [index, entry] of progress.actions.entries()) {
      if (entry.callId) {
        actionIndexByCallId.set(entry.callId, index);
      }
    }
  }
}

function updateProgressFromFunctionOutput(progress, payload, actionIndexByCallId) {
  const index = actionIndexByCallId.get(payload.call_id);
  if (index === undefined) {
    return;
  }

  const action = progress.actions[index];
  if (!action) {
    return;
  }

  const output = String(payload.output ?? "").trim();
  action.status = looksLikeToolFailure(output) ? "ERR" : "OK";
  action.outputPreview = summarizeToolOutput(output);
}

function summarizeFunctionCall(payload) {
  const args = safeParseJson(payload.arguments);
  const name = payload.name ?? "tool";
  if (name === "shell_command" || name === "exec_command") {
    const command = toSingleLine(args?.command ?? args?.cmd ?? payload.arguments ?? "");
    const workdir = args?.workdir ? ` @ ${toSingleLine(args.workdir)}` : "";
    return {
      callId: payload.call_id ?? null,
      toolName: name,
      status: "RUN",
      label: "shell",
      text: `${command}${workdir}`.trim()
    };
  }

  const summaryValue =
    args?.path ??
    args?.uri ??
    args?.command ??
    args?.prompt ??
    args?.pattern ??
    args?.q ??
    args?.location ??
    payload.arguments ??
    "";

  return {
    callId: payload.call_id ?? null,
    toolName: name,
    status: "RUN",
    label: name,
    text: toSingleLine(summaryValue || name)
  };
}

function summarizeToolOutput(output) {
  if (!output) {
    return "";
  }

  const exitMatch = output.match(/Process exited with code (\d+)/i);
  if (exitMatch) {
    return `exit ${exitMatch[1]}`;
  }

  const commandMatch = output.match(/Command:\s+([^\n]+)/i);
  if (commandMatch) {
    return toSingleLine(commandMatch[1]);
  }

  const singleLine = toSingleLine(output);
  return singleLine.length <= 120 ? singleLine : `${singleLine.slice(0, 117).trimEnd()}...`;
}

function looksLikeToolFailure(output) {
  const normalized = String(output ?? "").toLowerCase();
  return normalized.includes("execution error") || normalized.includes("failed") || normalized.includes("traceback");
}

function buildCodexFailureDetails(stdout, stderr, threadId, code, signal, finalMessage) {
  return [
    "Codex execution failed.",
    threadId ? `Session: ${threadId}` : null,
    code !== null ? `Exit code: ${code}` : null,
    signal ? `Signal: ${signal}` : null,
    stderr.trim() ? `stderr:\n${stderr.trim()}` : null,
    stdout.trim() && !finalMessage ? `stdout:\n${stdout.trim()}` : null
  ]
    .filter(Boolean)
    .join("\n\n");
}

function safeParseJson(input) {
  if (typeof input !== "string" || !input.trim()) {
    return null;
  }
  try {
    return JSON.parse(input);
  } catch {
    return null;
  }
}

function toSingleLine(input) {
  return String(input ?? "")
    .replace(/\s+/g, " ")
    .trim();
}

export function buildCodexArgs(
  config,
  scopeState,
  prompt,
  outputFile,
  attachmentContext = createEmptyAttachmentContext(),
  options = {}
) {
  const projectRoot = getScopeProjectRoot(config, scopeState);
  const hasThread = Boolean(scopeState.threadId);
  const args = ["exec", "--json", "--output-last-message", outputFile, "--skip-git-repo-check", "-C", projectRoot];
  const model = options.model || config.codexModel;
  const reasoningEffort = options.reasoningEffort || "";
  const sandboxMode = options.sandboxMode || config.codexExecMode;

  if (model) {
    args.push("-m", model);
  }

  if (reasoningEffort) {
    args.push("-c", `reasoning_effort="${reasoningEffort}"`);
  }

  switch (sandboxMode) {
    case "workspace-write":
      args.push("--full-auto");
      break;
    case "danger-full-access":
      args.push("--dangerously-bypass-approvals-and-sandbox");
      break;
    case "read-only":
    default:
      args.push("-s", "read-only");
      break;
  }

  if (hasThread) {
    args.push("resume");
    for (const imagePath of attachmentContext.imagePaths ?? []) {
      args.push("--image", imagePath);
    }
    args.push(scopeState.threadId, prompt);
  } else {
    for (const imagePath of attachmentContext.imagePaths ?? []) {
      args.push("--image", imagePath);
    }
    args.push(prompt);
  }

  return args;
}

export function chunkDiscordMessage(text, maxChars = 1900) {
  const normalized = text.trim() || "(empty response)";
  if (normalized.length <= maxChars) {
    return [normalized];
  }

  const chunks = [];
  let remaining = normalized;

  while (remaining.length > maxChars) {
    let slice = remaining.slice(0, maxChars);
    const splitIndex = Math.max(slice.lastIndexOf("\n"), slice.lastIndexOf(" "));
    if (splitIndex > maxChars * 0.5) {
      slice = slice.slice(0, splitIndex);
    }
    chunks.push(slice.trim());
    remaining = remaining.slice(slice.length).trim();
  }

  if (remaining) {
    chunks.push(remaining);
  }

  return chunks;
}

export function createScopeKey(interaction) {
  if (!interaction.guildId) {
    return `dm:${interaction.user.id}`;
  }
  return `guild:${interaction.guildId}:channel:${interaction.channelId}:user:${interaction.user.id}`;
}

export function listPathLikeLs(config, scopeState, input) {
  const targetInput = input?.trim() || ".";
  const resolved = resolveScopedProjectPath(config, scopeState, targetInput, {
    mustExist: true,
    allowFiles: true
  });
  if (!resolved.ok) {
    return resolved.error;
  }

  const targetPath = resolved.path;
  const entries = readDirectoryEntries(targetPath);
  const header = `${targetPath}:`;

  return [header, ...entries].join("\n");
}

export async function prepareAttachmentsForPrompt(config, scopeState, userId, attachments) {
  const normalizedAttachments = dedupeAttachments(attachments);
  if (normalizedAttachments.length === 0) {
    return createEmptyAttachmentContext();
  }

  const projectRoot = getScopeProjectRoot(config, scopeState);
  const uploadDir = path.posix.join(projectRoot, ".discord-uploads", userId);
  fs.mkdirSync(uploadDir, { recursive: true });

  const savedFiles = [];
  const imagePaths = [];

  for (const attachment of normalizedAttachments) {
    const size = Number.isFinite(attachment.size) ? attachment.size : 0;
    if (config.discordMaxAttachmentBytes > 0 && size > config.discordMaxAttachmentBytes) {
      throw new Error(
        `Attachment ${attachment.name} is too large (${size} bytes). Max allowed is ${config.discordMaxAttachmentBytes} bytes.`
      );
    }

    const savedFile = await downloadAttachment(uploadDir, attachment);
    const kind = classifyAttachment(savedFile.path, attachment.contentType);
    const metadata = {
      originalName: attachment.name ?? path.posix.basename(savedFile.path),
      path: savedFile.path,
      contentType: attachment.contentType ?? "unknown",
      size,
      kind
    };

    savedFiles.push(metadata);
    if (kind === "image") {
      imagePaths.push(savedFile.path);
    }
  }

  return { savedFiles, imagePaths };
}

export function isAllowedUser(config, userId) {
  if (!hasProjectAcl(config)) {
    return config.discordAllowedUserIds.length === 0 || config.discordAllowedUserIds.includes(userId);
  }

  return config.discordUserProjects.has(userId);
}

export function getAllowedProjectsForUser(config, userId) {
  return config.discordUserProjects.get(userId) ?? [];
}

export function hasAssignedProjects(config, userId) {
  return getAllowedProjectsForUser(config, userId).length > 0;
}

export function isAdminUser(config, userId) {
  return config.discordAdminUserIds.includes(userId);
}

export function getEffectiveUserSettings(config, userId) {
  const permission = normalizePermissionLevel(config.discordUserPermissions.get(userId) ?? "default");
  const reasoningValue = normalizeReasoningValue(config.discordUserReasoning.get(userId) ?? "");
  return {
    permission,
    sandboxMode: permission === "full-access" ? "danger-full-access" : config.codexExecMode,
    model: config.discordUserModels.get(userId) ?? "",
    reasoningEffort: reasoningValue === "default" ? "" : reasoningValue
  };
}

export function normalizePermissionLevel(input) {
  return input === "full-access" ? "full-access" : "default";
}

export function normalizeReasoningValue(input) {
  const value = String(input ?? "").trim().toLowerCase();
  if (value === "default") {
    return "default";
  }
  if (["low", "medium", "high"].includes(value)) {
    return value;
  }
  return "";
}

export function serializeKeyValueMap(map) {
  return [...map.entries()]
    .map(([key, value]) => `${key}:${value}`)
    .join(";");
}

export function readLatestRateLimits(config) {
  const sessionsRoot = path.posix.join(getCodexHome(config), "sessions");
  if (!fs.existsSync(sessionsRoot)) {
    return { ok: false, error: `Codex sessions directory not found: ${sessionsRoot}` };
  }

  const sessionFiles = collectSessionFiles(sessionsRoot)
    .sort((left, right) => {
      const leftTime = fs.statSync(left).mtimeMs;
      const rightTime = fs.statSync(right).mtimeMs;
      return rightTime - leftTime;
    });

  for (const filePath of sessionFiles.slice(0, 40)) {
    const lines = fs.readFileSync(filePath, "utf8").split(/\r?\n/).reverse();
    for (const line of lines) {
      const trimmed = line.trim();
      if (!trimmed) {
        continue;
      }
      try {
        const parsed = JSON.parse(trimmed);
        const payload = parsed.payload ?? {};
        if (parsed.type === "event_msg" && payload.type === "token_count" && payload.rate_limits) {
          return {
            ok: true,
            filePath,
            timestamp: parsed.timestamp ?? null,
            rateLimits: payload.rate_limits
          };
        }
      } catch {
        // Ignore malformed session lines.
      }
    }
  }

  return { ok: false, error: "No recent token_count event with rate-limit data was found." };
}

export function cloneUserProjects(userProjects) {
  const cloned = new Map();
  for (const [userId, projects] of userProjects.entries()) {
    cloned.set(userId, [...projects]);
  }
  return cloned;
}

export function serializeUserProjects(userProjects) {
  return [...userProjects.entries()]
    .map(([userId, projects]) => `${userId}:${projects.join("|")}`)
    .join(";");
}

export function normalizeDiscordUserId(input) {
  const trimmed = input.trim();
  if (!/^\d{5,30}$/.test(trimmed)) {
    return { ok: false, error: "Discord user id must be a numeric snowflake string." };
  }
  return { ok: true, userId: trimmed };
}

export function normalizeProjectName(input) {
  const validation = validateProjectName(input);
  return validation.ok
    ? { ok: true, projectName: validation.name }
    : validation;
}

export function log(level, message, extra = {}) {
  console.log(JSON.stringify({ level, message, ...extra }));
}

export function findLatestProjectSession(config, scopeState) {
  const codexHome = getCodexHome(config);
  const sessionsRoot = path.posix.join(codexHome, "sessions");
  if (!fs.existsSync(sessionsRoot)) {
    return { ok: false, error: `Codex sessions directory not found: ${sessionsRoot}` };
  }

  const indexEntries = loadSessionIndex(path.posix.join(codexHome, "session_index.jsonl"));
  const projectRoot = getScopeProjectRoot(config, scopeState);
  const sessionFiles = collectSessionFiles(sessionsRoot);
  const matches = [];

  for (const filePath of sessionFiles) {
    const meta = readSessionMeta(filePath);
    if (!meta?.id || !meta.cwd || !isPathWithinProject(meta.cwd, projectRoot)) {
      continue;
    }

    const indexed = indexEntries.get(meta.id) ?? {};
    matches.push({
      sessionId: meta.id,
      filePath,
      cwd: meta.cwd,
      updatedAt: indexed.updatedAt ?? new Date(fs.statSync(filePath).mtime).toISOString(),
      threadName: indexed.threadName ?? null,
      source: meta.source ?? null,
      originator: meta.originator ?? null
    });
  }

  if (matches.length === 0) {
    return { ok: false, error: `No Codex session found for current project ${projectRoot}` };
  }

  matches.sort((left, right) => {
    const leftTime = Date.parse(left.updatedAt) || 0;
    const rightTime = Date.parse(right.updatedAt) || 0;
    return rightTime - leftTime;
  });

  return {
    ok: true,
    session: {
      ...matches[0],
      recentMessages: readRecentSessionMessages(matches[0].filePath, 3)
    }
  };
}

function buildCodexPrompt(config, scopeState, promptText, projectRoot, attachmentContext = createEmptyAttachmentContext()) {
  const parts = [];

  if (config.codexMessagePrefix) {
    parts.push(config.codexMessagePrefix);
  }

  parts.push(`Project root: ${projectRoot}`);
  parts.push(`Current working directory inside the project: ${scopeState.cwd}`);
  parts.push("You must stay inside the project root. Do not read, write, or modify files outside it.");

  if (scopeState.cwd !== projectRoot) {
    parts.push(`When using relative paths or shell commands, treat ${scopeState.cwd} as the current directory inside the project.`);
  }

  if ((attachmentContext.savedFiles?.length ?? 0) > 0) {
    parts.push("");
    parts.push("Attached files were downloaded inside the current project:");
    for (const file of attachmentContext.savedFiles) {
      parts.push(`- ${file.path} (${file.kind}, original name: ${file.originalName})`);
    }
    parts.push("Use these exact paths as part of the task. Read relevant non-image files before answering.");
  }

  parts.push("");
  parts.push("User request:");
  parts.push(promptText);

  return parts.join("\n");
}

function createEmptyAttachmentContext() {
  return {
    savedFiles: [],
    imagePaths: []
  };
}

function buildCodexEnv(config) {
  return config.codexHome
    ? { ...process.env, CODEX_HOME: config.codexHome }
    : process.env;
}

function normalizeAbsolutePosix(input) {
  return path.posix.normalize(input);
}

function resolveStateDir(input) {
  return path.isAbsolute(input) ? input : path.join(process.cwd(), input);
}

function parseCsv(input) {
  return input
    .split(/[;,]/)
    .map((value) => value.trim())
    .filter(Boolean);
}

function parseBoolean(input) {
  const value = String(input ?? "").trim().toLowerCase();
  return value === "1" || value === "true" || value === "yes" || value === "on";
}

function parseProjectList(input) {
  return dedupePreserveOrder(
    String(input ?? "")
      .split(/[|,;]/)
      .map((value) => value.trim())
      .filter(Boolean)
  );
}

function parseUserProjects(input) {
  const result = new Map();

  for (const rawEntry of input.split(";")) {
    const entry = rawEntry.trim();
    if (!entry) {
      continue;
    }

    const match = entry.match(/^([^:=\s]+)\s*[:=]\s*(.*)$/);
    if (!match) {
      throw new Error(`Invalid DISCORD_USER_PROJECTS entry: ${entry}`);
    }

    const [, userId, projectsRaw] = match;
    const projects = projectsRaw
      .split(/[|,]/)
      .map((value) => value.trim())
      .filter(Boolean);

    const existing = result.get(userId) ?? [];
    result.set(userId, dedupePreserveOrder([...existing, ...projects]));
  }

  return result;
}

function parseKeyValueMap(input) {
  const result = new Map();

  for (const rawEntry of input.split(";")) {
    const entry = rawEntry.trim();
    if (!entry) {
      continue;
    }

    const match = entry.match(/^([^:=\s]+)\s*[:=]\s*(.*)$/);
    if (!match) {
      throw new Error(`Invalid key/value entry: ${entry}`);
    }

    const [, key, rawValue] = match;
    const value = rawValue.trim();
    if (!value) {
      continue;
    }
    result.set(key, value);
  }

  return result;
}

function getCodexHome(config) {
  return config.codexHome || path.posix.join(os.homedir(), ".codex");
}

function validateConfiguredProjects(userProjects) {
  for (const [userId, projects] of userProjects.entries()) {
    if (!userId) {
      throw new Error("DISCORD_USER_PROJECTS contains an empty user id.");
    }
    for (const projectName of projects) {
      const validation = validateProjectName(projectName);
      if (!validation.ok) {
        throw new Error(`Invalid project name for Discord user ${userId}: ${projectName}`);
      }
    }
  }
}

function validateProjectNames(projects, label) {
  for (const projectName of projects) {
    const validation = validateProjectName(projectName);
    if (!validation.ok) {
      throw new Error(`Invalid project name in ${label}: ${projectName}`);
    }
  }
}

function sameScopeState(left, right) {
  return (
    (left.projectName ?? null) === (right.projectName ?? null) &&
    (left.cwd ?? "") === (right.cwd ?? "") &&
    (left.threadId ?? null) === (right.threadId ?? null)
  );
}

function normalizeScopeCwd(input, fallback) {
  if (typeof input !== "string" || !input.trim()) {
    return fallback;
  }

  const normalized = path.posix.normalize(input);
  return path.posix.isAbsolute(normalized) ? normalized : fallback;
}

function selectDefaultProjectName(config, userId) {
  const allowedProjects = getAllowedProjectsForUser(config, userId);
  if (allowedProjects.length === 0) {
    return inferProjectNameFromPath(config, userId, config.defaultCwd);
  }

  const configuredDefault = inferProjectNameFromPath(config, userId, config.defaultCwd);
  if (configuredDefault && allowedProjects.includes(configuredDefault)) {
    return configuredDefault;
  }

  return allowedProjects[0];
}

function inferProjectNameFromPath(config, userId, candidatePath) {
  const normalized = normalizeScopeCwd(candidatePath, "");
  if (!normalized || !isWithinRoot(normalized, config.workspaceRoot)) {
    return null;
  }

  const relative = path.posix.relative(config.workspaceRoot, normalized);
  if (!relative || relative === "." || relative.startsWith("..")) {
    return null;
  }

  const topLevel = relative.split("/")[0];
  if (!topLevel || topLevel === "." || topLevel === "..") {
    return null;
  }

  if (!hasProjectAcl(config)) {
    return topLevel;
  }

  return getAllowedProjectsForUser(config, userId).includes(topLevel) ? topLevel : null;
}

function hasProjectAcl(config) {
  return config.discordUserProjects.size > 0;
}

function isProjectAllowedForUser(config, userId, projectName) {
  if (!hasProjectAcl(config)) {
    return true;
  }
  return getAllowedProjectsForUser(config, userId).includes(projectName);
}

function getProjectRoot(config, projectName) {
  return path.posix.join(config.workspaceRoot, projectName);
}

function getScopeProjectRoot(config, scopeState) {
  if (scopeState.projectName) {
    return getProjectRoot(config, scopeState.projectName);
  }

  const inferred = inferProjectNameFromPath(config, "", scopeState.cwd);
  return inferred ? getProjectRoot(config, inferred) : config.workspaceRoot;
}

function validateProjectName(name) {
  const trimmed = name.trim();
  if (!trimmed) {
    return { ok: false, error: "Project name cannot be empty." };
  }
  if (
    trimmed.includes("/") ||
    trimmed.includes("\\") ||
    trimmed.includes(":") ||
    trimmed.includes(";") ||
    trimmed.includes("|") ||
    trimmed.includes(",") ||
    trimmed === "." ||
    trimmed === ".."
  ) {
    return { ok: false, error: "Project name must be a single top-level directory name without config separators." };
  }
  return { ok: true, name: trimmed };
}

function resolvePathWithinRoot(rootDir, baseDir, input, options = {}) {
  const targetInput = input?.trim() || ".";
  const candidate = targetInput.startsWith("/")
    ? path.posix.normalize(targetInput)
    : path.posix.normalize(path.posix.join(baseDir, targetInput));

  if (!isWithinRoot(candidate, rootDir)) {
    return { ok: false, error: `Refusing to leave current project root ${rootDir}` };
  }

  if (options.mustExist) {
    try {
      const stat = fs.statSync(candidate);
      if (!options.allowFiles && !stat.isDirectory()) {
        return { ok: false, error: `${candidate} is not a directory` };
      }
    } catch {
      return {
        ok: false,
        error: options.allowFiles ? `Path does not exist: ${candidate}` : `Directory does not exist: ${candidate}`
      };
    }
  }

  return { ok: true, path: candidate };
}

function isWithinRoot(candidate, rootDir) {
  const normalizedCandidate = normalizeComparablePath(candidate);
  const normalizedRoot = normalizeComparablePath(rootDir);
  return normalizedCandidate === normalizedRoot || normalizedCandidate.startsWith(`${normalizedRoot}/`);
}

function isPathWithinProject(candidate, projectRoot) {
  return isWithinRoot(candidate, projectRoot);
}

function outputFilePath(filePath) {
  return filePath;
}

function dedupePreserveOrder(values) {
  return [...new Set(values)];
}

function dedupeAttachments(attachments) {
  const seen = new Set();
  const result = [];

  for (const attachment of attachments ?? []) {
    if (!attachment?.url) {
      continue;
    }
    const key = `${attachment.id ?? ""}:${attachment.url}`;
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    result.push(attachment);
  }

  return result;
}

async function downloadAttachment(uploadDir, attachment) {
  const safeName = sanitizeFilename(attachment.name ?? "attachment");
  const fileName = `${Date.now()}-${Math.random().toString(16).slice(2, 10)}-${safeName}`;
  const filePath = path.posix.join(uploadDir, fileName);

  const response = await fetch(attachment.url);
  if (!response.ok) {
    throw new Error(`Failed to download attachment ${attachment.name}: HTTP ${response.status}`);
  }

  const arrayBuffer = await response.arrayBuffer();
  fs.writeFileSync(filePath, Buffer.from(arrayBuffer));
  return { path: filePath };
}

function sanitizeFilename(input) {
  const base = path.posix.basename(input);
  const sanitized = base.replace(/[^A-Za-z0-9._-]/g, "_");
  return sanitized || "attachment";
}

function classifyAttachment(filePath, contentType) {
  const normalizedType = (contentType ?? "").toLowerCase();
  if (normalizedType.startsWith("image/")) {
    return "image";
  }

  const extension = path.posix.extname(filePath).toLowerCase();
  if ([".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp"].includes(extension)) {
    return "image";
  }
  if (
    normalizedType.startsWith("text/") ||
    [
      ".txt",
      ".md",
      ".json",
      ".yaml",
      ".yml",
      ".toml",
      ".xml",
      ".csv",
      ".log",
      ".js",
      ".jsx",
      ".ts",
      ".tsx",
      ".py",
      ".sh",
      ".rb",
      ".go",
      ".rs",
      ".java",
      ".kt",
      ".php",
      ".c",
      ".cc",
      ".cpp",
      ".h",
      ".hpp",
      ".cs",
      ".sql"
    ].includes(extension)
  ) {
    return "text";
  }

  return "file";
}

function listWorkspaceProjects(config) {
  const entries = fs.readdirSync(config.workspaceRoot, { withFileTypes: true });
  return entries
    .filter((entry) => entry.isDirectory())
    .map((entry) => entry.name)
    .sort();
}

function loadSessionIndex(filePath) {
  const entries = new Map();
  if (!fs.existsSync(filePath)) {
    return entries;
  }

  for (const line of fs.readFileSync(filePath, "utf8").split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed) {
      continue;
    }
    try {
      const parsed = JSON.parse(trimmed);
      if (parsed.id) {
        entries.set(parsed.id, {
          threadName: parsed.thread_name ?? null,
          updatedAt: parsed.updated_at ?? null
        });
      }
    } catch {
      // Ignore malformed index lines.
    }
  }

  return entries;
}

function collectSessionFiles(rootDir) {
  const files = [];
  const queue = [rootDir];

  while (queue.length > 0) {
    const current = queue.pop();
    for (const entry of fs.readdirSync(current, { withFileTypes: true })) {
      const fullPath = path.posix.join(current, entry.name);
      if (entry.isDirectory()) {
        queue.push(fullPath);
        continue;
      }
      if (entry.isFile() && fullPath.endsWith(".jsonl")) {
        files.push(fullPath);
      }
    }
  }

  return files;
}

function findSessionFileById(config, threadId) {
  if (!threadId) {
    return null;
  }

  const sessionsRoot = path.posix.join(getCodexHome(config), "sessions");
  if (!fs.existsSync(sessionsRoot)) {
    return null;
  }

  return collectSessionFiles(sessionsRoot).find((filePath) => filePath.includes(threadId)) ?? null;
}

function readSessionMeta(filePath) {
  try {
    const handle = fs.openSync(filePath, "r");
    try {
      const firstLine = readFirstLine(handle);
      if (!firstLine) {
        return null;
      }
      const parsed = JSON.parse(firstLine);
      return parsed?.type === "session_meta" ? parsed.payload ?? null : null;
    } finally {
      fs.closeSync(handle);
    }
  } catch {
    return null;
  }
}

function readFirstLine(handle) {
  const chunks = [];
  const buffer = Buffer.alloc(4096);
  let position = 0;

  while (position < 1024 * 1024) {
    const size = fs.readSync(handle, buffer, 0, buffer.length, position);
    if (size <= 0) {
      break;
    }

    const slice = buffer.subarray(0, size);
    const newlineIndex = slice.indexOf(0x0a);
    if (newlineIndex >= 0) {
      chunks.push(Buffer.from(slice.subarray(0, newlineIndex)));
      break;
    }

    chunks.push(Buffer.from(slice));
    position += size;
  }

  return Buffer.concat(chunks).toString("utf8").replace(/\r$/, "");
}

function readRecentSessionMessages(filePath, count) {
  const messages = [];

  for (const line of fs.readFileSync(filePath, "utf8").split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed) {
      continue;
    }
    try {
      const parsed = JSON.parse(trimmed);
      if (parsed.type !== "response_item" || parsed.payload?.type !== "message") {
        continue;
      }

      const role = parsed.payload.role;
      if (role !== "user" && role !== "assistant") {
        continue;
      }

      const text = extractMessageText(parsed.payload.content ?? []);
      if (!text || text.startsWith("<environment_context>")) {
        continue;
      }

      messages.push({
        role,
        phase: parsed.payload.phase ?? null,
        timestamp: parsed.timestamp ?? null,
        text
      });
    } catch {
      // Ignore malformed session lines.
    }
  }

  return messages.slice(-count);
}

function extractMessageText(content) {
  const parts = [];
  for (const item of content) {
    if (item?.type === "input_text" || item?.type === "output_text") {
      const text = String(item.text ?? "").trim();
      if (text) {
        parts.push(text);
      }
    }
  }
  return parts.join("\n\n").trim();
}

function normalizeComparablePath(input) {
  const raw = String(input ?? "").trim();
  if (!raw) {
    return "";
  }

  const windowsMatch = raw.match(/^([A-Za-z]):[\\/](.*)$/);
  if (windowsMatch) {
    const [, drive, rest] = windowsMatch;
    const normalizedRest = rest.replace(/\\/g, "/").replace(/^\/+/, "");
    return path.posix.normalize(`/mnt/${drive.toLowerCase()}/${normalizedRest}`);
  }

  return path.posix.normalize(raw.replace(/\\/g, "/"));
}

function isExistingDirectory(targetPath) {
  try {
    return fs.statSync(targetPath).isDirectory();
  } catch {
    return false;
  }
}

function loadJson(filePath, fallback) {
  try {
    if (!fs.existsSync(filePath)) {
      return fallback;
    }
    return JSON.parse(fs.readFileSync(filePath, "utf8"));
  } catch {
    return fallback;
  }
}

function readDirectoryEntries(targetPath) {
  const stats = fs.statSync(targetPath);
  const names = [".", ".."];

  if (stats.isDirectory()) {
    const children = fs.readdirSync(targetPath, { withFileTypes: true }).map((entry) => entry.name).sort();
    names.push(...children);
  } else {
    names.push(path.posix.basename(targetPath));
  }

  return names.map((name) => formatDirentLine(targetPath, name));
}

function formatDirentLine(targetPath, name) {
  const fullPath =
    name === "."
      ? targetPath
      : name === ".."
        ? path.posix.dirname(targetPath)
        : path.posix.join(targetPath, name);

  let stats;
  try {
    stats = fs.lstatSync(fullPath);
  } catch {
    return `?????????? ? ? ? ? ${name}`;
  }

  const typeChar = stats.isDirectory()
    ? "d"
    : stats.isSymbolicLink()
      ? "l"
      : "-";

  const perms = [
    stats.mode & 0o400 ? "r" : "-",
    stats.mode & 0o200 ? "w" : "-",
    stats.mode & 0o100 ? "x" : "-",
    stats.mode & 0o040 ? "r" : "-",
    stats.mode & 0o020 ? "w" : "-",
    stats.mode & 0o010 ? "x" : "-",
    stats.mode & 0o004 ? "r" : "-",
    stats.mode & 0o002 ? "w" : "-",
    stats.mode & 0o001 ? "x" : "-"
  ].join("");

  const nlink = String(stats.nlink).padStart(2, " ");
  const uid = String(stats.uid).padStart(5, " ");
  const gid = String(stats.gid).padStart(5, " ");
  const size = String(stats.size).padStart(8, " ");
  const mtime = formatMtime(stats.mtime);

  return `${typeChar}${perms} ${nlink} ${uid} ${gid} ${size} ${mtime} ${name}`;
}

function formatMtime(value) {
  const date = new Date(value);
  const month = date.toLocaleString("en-US", { month: "short" });
  const day = String(date.getDate()).padStart(2, " ");
  const hours = String(date.getHours()).padStart(2, "0");
  const minutes = String(date.getMinutes()).padStart(2, "0");
  return `${month} ${day} ${hours}:${minutes}`;
}
