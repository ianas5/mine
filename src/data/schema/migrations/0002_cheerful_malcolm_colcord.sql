CREATE TABLE `sets` (
	`id` text PRIMARY KEY NOT NULL,
	`workout_exercise_id` text NOT NULL,
	`position` integer NOT NULL,
	`weight_kg` real DEFAULT 0 NOT NULL,
	`reps` integer DEFAULT 0 NOT NULL,
	`rpe` real,
	`rir` integer,
	`is_warmup` integer DEFAULT 0 NOT NULL,
	`notes` text,
	FOREIGN KEY (`workout_exercise_id`) REFERENCES `workout_exercises`(`id`) ON UPDATE no action ON DELETE cascade,
	CONSTRAINT "sets_rpe_range" CHECK("sets"."rpe" IS NULL OR ("sets"."rpe" >= 0 AND "sets"."rpe" <= 10)),
	CONSTRAINT "sets_rir_range" CHECK("sets"."rir" IS NULL OR ("sets"."rir" >= 0 AND "sets"."rir" <= 10))
);
--> statement-breakpoint
CREATE INDEX `sets_workout_exercise` ON `sets` (`workout_exercise_id`);--> statement-breakpoint
CREATE TABLE `workout_exercises` (
	`id` text PRIMARY KEY NOT NULL,
	`workout_id` text NOT NULL,
	`exercise_id` text NOT NULL,
	`position` integer NOT NULL,
	`unilateral_counting` text DEFAULT 'none' NOT NULL,
	`notes` text,
	FOREIGN KEY (`workout_id`) REFERENCES `workouts`(`id`) ON UPDATE no action ON DELETE cascade,
	FOREIGN KEY (`exercise_id`) REFERENCES `exercises`(`id`) ON UPDATE no action ON DELETE restrict,
	CONSTRAINT "workout_exercises_counting" CHECK(unilateral_counting = 'none' OR unilateral_counting = 'single_doubled' OR unilateral_counting = 'per_side')
);
--> statement-breakpoint
CREATE INDEX `workout_exercises_workout` ON `workout_exercises` (`workout_id`);--> statement-breakpoint
CREATE INDEX `workout_exercises_exercise` ON `workout_exercises` (`exercise_id`);--> statement-breakpoint
CREATE TABLE `workouts` (
	`id` text PRIMARY KEY NOT NULL,
	`date` text NOT NULL,
	`name` text NOT NULL,
	`template_id` text,
	`started_at` integer,
	`ended_at` integer,
	`notes` text,
	`created_at` integer NOT NULL,
	`updated_at` integer NOT NULL
);
--> statement-breakpoint
CREATE INDEX `workouts_date` ON `workouts` (`date`);