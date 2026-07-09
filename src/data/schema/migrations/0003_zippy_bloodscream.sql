CREATE TABLE `workout_drafts` (
	`id` integer PRIMARY KEY NOT NULL,
	`payload` text NOT NULL,
	`updated_at` integer NOT NULL,
	CONSTRAINT "workout_drafts_single_row" CHECK("workout_drafts"."id" = 1)
);
