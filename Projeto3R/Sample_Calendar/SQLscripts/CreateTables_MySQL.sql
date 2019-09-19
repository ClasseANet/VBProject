-- MySQL-Administrator dump 1.3
--
-- ------------------------------------------------------
-- Server version	4.1.9-nt-max


SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT;
SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS;
SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION;
SET NAMES utf8;

/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='ANSI_QUOTES,NO_AUTO_VALUE_ON_ZERO' */;



DROP TABLE IF EXISTS "calendarevents";
CREATE TABLE "calendarevents" (
  "EventID" int(10) unsigned NOT NULL auto_increment,
  "StartDateTime" datetime NOT NULL default '0000-00-00 00:00:00',
  "EndDateTime" datetime NOT NULL default '0000-00-00 00:00:00',
  "RecurrenceState" int(11) NOT NULL default '0',
  "Subject" text NOT NULL,
  "Location" text NOT NULL,
  "Body" text NOT NULL,
  "BusyStatus" int(11) NOT NULL default '0',
  "ImportanceLevel" int(11) NOT NULL default '0',
  "LabelID" int(11) NOT NULL default '0',
  "ScheduleID" int(11) NOT NULL default '0',
  "RecurrencePatternID" int(11) NOT NULL default '0',
  "IsRecurrenceExceptionDeleted" int(11) NOT NULL default '0',
  "RExceptionStartTimeOrig" datetime NOT NULL default '0000-00-00 00:00:00',
  "RExceptionEndTimeOrig" datetime NOT NULL default '0000-00-00 00:00:00',
  "IsAllDayEvent" int(11) NOT NULL default '0',
  "IsMeeting" int(11) NOT NULL default '0',
  "IsPrivate" int(11) NOT NULL default '0',
  "IsReminder" int(11) NOT NULL default '0',
  "ReminderMinutesBeforeStart" int(11) NOT NULL default '0',
  "RemainderSoundFile" text NOT NULL,
  "CustomPropertiesXMLData" text NOT NULL,
  "CustomIconsIDs" text NOT NULL,
  "Created" datetime NOT NULL default '0000-00-00 00:00:00',
  "Modified" datetime NOT NULL default '0000-00-00 00:00:00',
  PRIMARY KEY  ("EventID")
) ENGINE=MyISAM DEFAULT CHARSET=latin1 ROW_FORMAT=FIXED;

DROP TABLE IF EXISTS "calendarrecurrencepatterns";
CREATE TABLE "calendarrecurrencepatterns" (
  "RecurrencePatternID" int(11) unsigned NOT NULL auto_increment,
  "MasterEventID" int(11) default NULL,
  "PatternStartDate" datetime default NULL,
  "PatternEndMethod" int(11) default NULL,
  "PatternEndDate" datetime default NULL,
  "PatternEndAfterOccurrences" int(11) default NULL,
  "EventStartTime" datetime default NULL,
  "EventDuration" int(11) default NULL,
  "OptionsData1" int(11) default NULL,
  "OptionsData2" int(11) default NULL,
  "OptionsData3" int(11) default NULL,
  "OptionsData4" int(11) default NULL,
  "CustomPropertiesXMLData" mediumtext,
  PRIMARY KEY  ("RecurrencePatternID")
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT;
SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS;
SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
