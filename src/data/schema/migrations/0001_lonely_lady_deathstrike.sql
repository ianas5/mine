CREATE TABLE `exercises` (
	`id` text PRIMARY KEY NOT NULL,
	`name` text NOT NULL,
	`primary_muscle_group` text NOT NULL,
	`secondary_muscle_groups` text DEFAULT '[]' NOT NULL,
	`load_type` text DEFAULT 'external' NOT NULL,
	`default_unilateral` integer DEFAULT 0 NOT NULL,
	`is_custom` integer DEFAULT 0 NOT NULL,
	`is_archived` integer DEFAULT 0 NOT NULL,
	`notes` text,
	`created_at` integer NOT NULL,
	`updated_at` integer NOT NULL,
	CONSTRAINT "exercises_primary_group" CHECK(primary_muscle_group = 'chest' OR primary_muscle_group = 'shoulders' OR primary_muscle_group = 'back' OR primary_muscle_group = 'biceps' OR primary_muscle_group = 'triceps' OR primary_muscle_group = 'forearms' OR primary_muscle_group = 'core' OR primary_muscle_group = 'glutes' OR primary_muscle_group = 'quads' OR primary_muscle_group = 'hamstrings' OR primary_muscle_group = 'calves' OR primary_muscle_group = 'other'),
	CONSTRAINT "exercises_load_type" CHECK(load_type = 'external' OR load_type = 'bodyweight' OR load_type = 'bodyweight_plus' OR load_type = 'assisted' OR load_type = 'timed')
);
--> statement-breakpoint
CREATE UNIQUE INDEX `exercises_name_nocase` ON `exercises` ("name" COLLATE NOCASE);