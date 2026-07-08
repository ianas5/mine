CREATE TABLE `settings` (
	`id` integer PRIMARY KEY NOT NULL,
	`weekly_workout_target` integer DEFAULT 4 NOT NULL,
	`default_bodyweight_kg` real,
	`height_cm` real,
	`water_cup_ml` integer DEFAULT 250 NOT NULL,
	`created_at` integer NOT NULL,
	`updated_at` integer NOT NULL,
	CONSTRAINT "settings_single_row" CHECK("settings"."id" = 1)
);
