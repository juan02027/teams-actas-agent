CREATE TABLE `agentSettings` (
	`id` int AUTO_INCREMENT NOT NULL,
	`companyName` varchar(180) NOT NULL DEFAULT 'Toda la empresa',
	`recordByDefault` int NOT NULL DEFAULT 0,
	`processByDefault` int NOT NULL DEFAULT 1,
	`requireReview` int NOT NULL DEFAULT 1,
	`destination` varchar(255) NOT NULL DEFAULT 'SharePoint / Actas de reuniones',
	`aiModel` varchar(120) NOT NULL DEFAULT 'gpt-5-mini',
	`updatedAt` timestamp NOT NULL DEFAULT (now()) ON UPDATE CURRENT_TIMESTAMP,
	CONSTRAINT `agentSettings_id` PRIMARY KEY(`id`)
);
--> statement-breakpoint
CREATE TABLE `commitments` (
	`id` int AUTO_INCREMENT NOT NULL,
	`meetingId` int NOT NULL,
	`personName` varchar(180) NOT NULL,
	`personEmail` varchar(320),
	`action` text NOT NULL,
	`dueDate` varchar(64),
	`status` enum('open','in_progress','done','blocked') NOT NULL DEFAULT 'open',
	`createdAt` timestamp NOT NULL DEFAULT (now()),
	CONSTRAINT `commitments_id` PRIMARY KEY(`id`)
);
--> statement-breakpoint
CREATE TABLE `meetingDocuments` (
	`id` int AUTO_INCREMENT NOT NULL,
	`meetingId` int NOT NULL,
	`kind` enum('minutes','commitments') NOT NULL,
	`format` enum('docx','pdf') NOT NULL,
	`fileName` varchar(255) NOT NULL,
	`storageKey` text,
	`storageUrl` text,
	`generatedAt` timestamp NOT NULL DEFAULT (now()),
	CONSTRAINT `meetingDocuments_id` PRIMARY KEY(`id`)
);
--> statement-breakpoint
CREATE TABLE `meetings` (
	`id` int AUTO_INCREMENT NOT NULL,
	`graphMeetingId` varchar(180) NOT NULL,
	`title` varchar(255) NOT NULL,
	`organizerName` varchar(180) NOT NULL,
	`organizerEmail` varchar(320),
	`scheduledAt` timestamp NOT NULL,
	`durationMinutes` int NOT NULL DEFAULT 0,
	`attendeesCount` int NOT NULL DEFAULT 0,
	`recordingEnabled` int NOT NULL DEFAULT 0,
	`processingEnabled` int NOT NULL DEFAULT 1,
	`status` enum('scheduled','recording','processing','ready','review','error') NOT NULL DEFAULT 'scheduled',
	`transcriptUrl` text,
	`recordingUrl` text,
	`createdAt` timestamp NOT NULL DEFAULT (now()),
	`updatedAt` timestamp NOT NULL DEFAULT (now()) ON UPDATE CURRENT_TIMESTAMP,
	CONSTRAINT `meetings_id` PRIMARY KEY(`id`),
	CONSTRAINT `meetings_graphMeetingId_unique` UNIQUE(`graphMeetingId`)
);
--> statement-breakpoint
CREATE INDEX `commitments_meeting_idx` ON `commitments` (`meetingId`);--> statement-breakpoint
CREATE INDEX `meeting_documents_meeting_idx` ON `meetingDocuments` (`meetingId`);--> statement-breakpoint
CREATE INDEX `meetings_graph_meeting_idx` ON `meetings` (`graphMeetingId`);