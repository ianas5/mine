CREATE TABLE `foods` (
	`id` text PRIMARY KEY NOT NULL,
	`name` text NOT NULL,
	`serving_amount` real NOT NULL,
	`serving_unit` text NOT NULL,
	`kcal` integer NOT NULL,
	`protein_g` real NOT NULL,
	`carb_g` real NOT NULL,
	`fat_g` real NOT NULL,
	`is_quick_meal` integer DEFAULT 0 NOT NULL,
	`is_custom` integer DEFAULT 1 NOT NULL,
	`is_archived` integer DEFAULT 0 NOT NULL,
	`created_at` integer NOT NULL,
	`updated_at` integer NOT NULL,
	CONSTRAINT "foods_serving_unit" CHECK(serving_unit = 'g' OR serving_unit = 'ml' OR serving_unit = 'piece' OR serving_unit = 'scoop' OR serving_unit = 'cup' OR serving_unit = 'serving')
);
--> statement-breakpoint
CREATE TABLE `meal_entries` (
	`id` text PRIMARY KEY NOT NULL,
	`date` text NOT NULL,
	`slot` text,
	`food_id` text,
	`food_name` text NOT NULL,
	`logged_amount` real NOT NULL,
	`logged_unit` text NOT NULL,
	`kcal` integer NOT NULL,
	`protein_g` real NOT NULL,
	`carb_g` real NOT NULL,
	`fat_g` real NOT NULL,
	`logged_at` integer NOT NULL,
	FOREIGN KEY (`food_id`) REFERENCES `foods`(`id`) ON UPDATE no action ON DELETE set null,
	CONSTRAINT "meal_entries_slot" CHECK("meal_entries"."slot" IS NULL OR (slot = 'breakfast' OR slot = 'lunch' OR slot = 'dinner' OR slot = 'snacks'))
);
--> statement-breakpoint
CREATE INDEX `meal_entries_date` ON `meal_entries` (`date`);--> statement-breakpoint
CREATE INDEX `meal_entries_food` ON `meal_entries` (`food_id`);