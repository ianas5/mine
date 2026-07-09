CREATE TABLE `progress_photos` (
	`id` text PRIMARY KEY NOT NULL,
	`date` text NOT NULL,
	`angle` text NOT NULL,
	`file_name` text NOT NULL,
	`width` integer,
	`height` integer,
	`notes` text,
	`created_at` integer NOT NULL,
	CONSTRAINT "progress_photos_angle" CHECK(angle = 'front' OR angle = 'side' OR angle = 'back')
);
--> statement-breakpoint
CREATE UNIQUE INDEX `progress_photos_file_name_unique` ON `progress_photos` (`file_name`);--> statement-breakpoint
CREATE INDEX `progress_photos_date` ON `progress_photos` (`date`);