CREATE TABLE `nutrition_targets` (
	`id` text PRIMARY KEY NOT NULL,
	`effective_from` text NOT NULL,
	`kcal` integer NOT NULL,
	`protein_g` real NOT NULL,
	`carb_g` real NOT NULL,
	`fat_g` real NOT NULL,
	`water_ml` integer,
	`created_at` integer NOT NULL,
	`updated_at` integer NOT NULL
);
--> statement-breakpoint
CREATE UNIQUE INDEX `nutrition_targets_effective_from_unique` ON `nutrition_targets` (`effective_from`);--> statement-breakpoint
CREATE TABLE `water_days` (
	`date` text PRIMARY KEY NOT NULL,
	`ml` integer DEFAULT 0 NOT NULL,
	`updated_at` integer NOT NULL
);
