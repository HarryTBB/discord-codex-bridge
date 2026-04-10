import dotenv from "dotenv";

import { loadConfig, log, registerSlashCommands } from "./shared.mjs";

dotenv.config();

const config = loadConfig();

const count = await registerSlashCommands(config);
log("info", `Registered ${count} Discord slash commands`, {
  applicationId: config.discordApplicationId,
  guildId: config.discordGuildId
});
