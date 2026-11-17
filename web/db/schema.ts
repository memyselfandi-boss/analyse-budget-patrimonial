import { sqliteTable, text, integer, index } from "drizzle-orm/sqlite-core";
export const clientFiles = sqliteTable("client_files", {
 id: text("id").primaryKey(), owner: text("owner").notNull(),
 name: text("name").notNull(), payload: text("payload").notNull(),
 version: integer("version").notNull(), createdAt: text("created_at").notNull(), updatedAt: text("updated_at").notNull(),
}, t => [index("client_files_owner_updated").on(t.owner,t.updatedAt)]);
