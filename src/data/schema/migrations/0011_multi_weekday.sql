ALTER TABLE `templates` ADD `weekdays` text DEFAULT '[]' NOT NULL;--> statement-breakpoint
UPDATE `templates` SET `weekdays` = '[' || `weekday` || ']' WHERE `weekday` IS NOT NULL;