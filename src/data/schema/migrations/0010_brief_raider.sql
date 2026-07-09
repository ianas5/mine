CREATE TABLE `phases` (
	`id` text PRIMARY KEY NOT NULL,
	`name` text NOT NULL,
	`type` text NOT NULL,
	`start_date` text NOT NULL,
	`end_date` text,
	`notes` text,
	`created_at` integer NOT NULL,
	`updated_at` integer NOT NULL,
	CONSTRAINT "phases_type" CHECK(type = 'cutting' OR type = 'recomp' OR type = 'lean_bulk' OR type = 'maintenance' OR type = 'custom'),
	CONSTRAINT "phases_dates" CHECK("phases"."end_date" IS NULL OR "phases"."end_date" >= "phases"."start_date")
);
--> statement-breakpoint
CREATE INDEX `phases_start` ON `phases` (`start_date`);