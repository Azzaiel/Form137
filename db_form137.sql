-- phpMyAdmin SQL Dump
-- version 4.0.4
-- http://www.phpmyadmin.net
--
-- Host: localhost
-- Generation Time: Jan 16, 2014 at 12:42 AM
-- Server version: 5.6.12-log
-- PHP Version: 5.4.16

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;

--
-- Database: `db_form137`
--
CREATE DATABASE IF NOT EXISTS `db_form137` DEFAULT CHARACTER SET latin1 COLLATE latin1_swedish_ci;
USE `db_form137`;

-- --------------------------------------------------------

--
-- Table structure for table `tbl_attendance`
--

CREATE TABLE IF NOT EXISTS `tbl_attendance` (
  `SY` varchar(200) NOT NULL,
  `ID` varchar(200) NOT NULL,
  `section_name` varchar(200) NOT NULL,
  `no_school_days` varchar(200) NOT NULL,
  `no_days_absent` varchar(200) NOT NULL,
  `causes_of_absences` varchar(500) NOT NULL,
  `no_days_tardiness` varchar(200) NOT NULL,
  `causes_of_tardiness` varchar(500) NOT NULL,
  `no_days_present` varchar(200) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tbl_attendance`
--

INSERT INTO `tbl_attendance` (`SY`, `ID`, `section_name`, `no_school_days`, `no_days_absent`, `causes_of_absences`, `no_days_tardiness`, `causes_of_tardiness`, `no_days_present`) VALUES
('2013-2014', '109637111111', 'Apple', '180', '0', '  ', '0', '  ', '180');

-- --------------------------------------------------------

--
-- Table structure for table `tbl_character_grade`
--

CREATE TABLE IF NOT EXISTS `tbl_character_grade` (
  `No` int(100) NOT NULL AUTO_INCREMENT,
  `SY` varchar(200) NOT NULL,
  `ID` varchar(200) NOT NULL,
  `section_name` varchar(100) NOT NULL,
  `Period` varchar(200) NOT NULL,
  `Honesty` varchar(200) NOT NULL,
  `Courtesy` varchar(200) NOT NULL,
  `Helpfulness_and_Cooperation` varchar(200) NOT NULL,
  `Resourcefulness_and_Creativity` varchar(200) NOT NULL,
  `Consideration_for_Others` varchar(200) NOT NULL,
  `Sportsmanship` varchar(200) NOT NULL,
  `Obedience` varchar(200) NOT NULL,
  `Self_Reliance` varchar(200) NOT NULL,
  `Industry` varchar(200) NOT NULL,
  `Cleanliness_and_Orderliness` varchar(200) NOT NULL,
  `Promptness_and_Punctuality` varchar(200) NOT NULL,
  `Sense_of_Responsibility` varchar(200) NOT NULL,
  `Love_of_God` varchar(200) NOT NULL,
  `Patriotism_and_Love_of_Country` varchar(200) NOT NULL,
  PRIMARY KEY (`No`)
) ENGINE=InnoDB  DEFAULT CHARSET=latin1 AUTO_INCREMENT=11 ;

--
-- Dumping data for table `tbl_character_grade`
--

INSERT INTO `tbl_character_grade` (`No`, `SY`, `ID`, `section_name`, `Period`, `Honesty`, `Courtesy`, `Helpfulness_and_Cooperation`, `Resourcefulness_and_Creativity`, `Consideration_for_Others`, `Sportsmanship`, `Obedience`, `Self_Reliance`, `Industry`, `Cleanliness_and_Orderliness`, `Promptness_and_Punctuality`, `Sense_of_Responsibility`, `Love_of_God`, `Patriotism_and_Love_of_Country`) VALUES
(1, '2013-2014', '109637111111', 'Apple', '1st Grading', 'A', 'B', 'A', 'B', 'A', 'B', 'A', 'B', 'A', 'B', 'A', 'B', 'A', 'B'),
(2, '2013-2014', '109637222222', 'Apple', '1st Grading', 'A', 'A', 'A', 'A', 'A', 'A', 'A', 'A', 'A', 'A', 'A', 'A', 'A', 'A'),
(3, '2013-2014', '109637111111', 'Apple', '2nd Grading', 'A', 'A', 'A', 'A', 'A', 'A', 'A', 'A', 'A', 'A', 'A', 'A', 'A', 'A'),
(4, '2013-2014', '109637222222', 'Apple', '2nd Grading', 'B', 'B', 'B', 'B', 'B', 'B', 'B', 'B', 'B', 'B', 'B', 'B', 'B', 'B'),
(5, '2013-2014', '109637111111', 'Apple', 'Final', 'A', 'B', 'A', 'B', 'A', 'B', 'A', 'B', 'A', 'B', 'A', 'B', 'A', 'B'),
(6, '2013-2014', '109637222222', 'Apple', 'Final', '', '', '', '', '', '', '', '', '', '', '', '', '', ''),
(7, '2013-2014', '109637111111', 'Apple', '3rd Grading', 'B', 'B', 'B', 'B', 'B', 'B', 'B', 'B', 'B', 'B', 'B', 'B', 'B', 'B'),
(8, '2013-2014', '109637222222', 'Apple', '3rd Grading', '', '', '', '', '', '', '', '', '', '', '', '', '', ''),
(9, '2013-2014', '109637111111', 'Apple', '4th Grading', 'A', 'A', 'A', 'A', 'A', 'A', 'A', 'A', 'A', 'A', 'A', 'A', 'A', 'A'),
(10, '2013-2014', '109637222222', 'Apple', '4th Grading', '', '', '', '', '', '', '', '', '', '', '', '', '', '');

-- --------------------------------------------------------

--
-- Table structure for table `tbl_grade`
--

CREATE TABLE IF NOT EXISTS `tbl_grade` (
  `grd_id` int(11) NOT NULL AUTO_INCREMENT,
  `SY` varchar(200) NOT NULL,
  `ID` varchar(200) NOT NULL,
  `section_name` varchar(200) NOT NULL,
  `subject_code` varchar(200) NOT NULL,
  `period` varchar(200) NOT NULL,
  `grade` varchar(100) NOT NULL,
  `remark` varchar(100) NOT NULL,
  PRIMARY KEY (`grd_id`)
) ENGINE=InnoDB  DEFAULT CHARSET=latin1 AUTO_INCREMENT=21 ;

--
-- Dumping data for table `tbl_grade`
--

INSERT INTO `tbl_grade` (`grd_id`, `SY`, `ID`, `section_name`, `subject_code`, `period`, `grade`, `remark`) VALUES
(1, '2013-2014', '109637111111', 'Apple', 'Eng', '1st Grading', '90', 'A'),
(2, '2013-2014', '109637111111', 'Apple', 'Eng', '2nd Grading', '90', 'A'),
(3, '2013-2014', '109637111111', 'Apple', 'Eng', '3rd Grading', '90', 'A'),
(4, '2013-2014', '109637111111', 'Apple', 'Eng', '4th Grading', '100', 'A'),
(5, '2013-2014', '109637111111', 'Apple', 'Eng', 'Final', '92.5', 'A'),
(6, '2013-2014', '109637222222', 'Apple', 'Eng', '1st Grading', '90', 'A'),
(7, '2013-2014', '109637222222', 'Apple', 'Eng', '2nd Grading', '90', 'A'),
(8, '2013-2014', '109637222222', 'Apple', 'Eng', '3rd Grading', '90', 'A'),
(9, '2013-2014', '109637222222', 'Apple', 'Eng', '4th Grading', '80', 'AP'),
(10, '2013-2014', '109637222222', 'Apple', 'Eng', 'Final', '87.5', 'P'),
(11, '2013-2014', '109637111111', 'Apple', 'Mat', '1st Grading', '90', 'A'),
(12, '2013-2014', '109637111111', 'Apple', 'Mat', '2nd Grading', '90', 'A'),
(13, '2013-2014', '109637111111', 'Apple', 'Mat', '3rd Grading', '90', 'A'),
(14, '2013-2014', '109637111111', 'Apple', 'Mat', '4th Grading', '85', 'P'),
(15, '2013-2014', '109637111111', 'Apple', 'Mat', 'Final', '88.75', 'P'),
(16, '2013-2014', '109637222222', 'Apple', 'Mat', '1st Grading', '90', 'A'),
(17, '2013-2014', '109637222222', 'Apple', 'Mat', '2nd Grading', '90', 'A'),
(18, '2013-2014', '109637222222', 'Apple', 'Mat', '3rd Grading', '90', 'A'),
(19, '2013-2014', '109637222222', 'Apple', 'Mat', '4th Grading', '95', 'A'),
(20, '2013-2014', '109637222222', 'Apple', 'Mat', 'Final', '91.25', 'A');

-- --------------------------------------------------------

--
-- Table structure for table `tbl_level`
--

CREATE TABLE IF NOT EXISTS `tbl_level` (
  `SY` varchar(100) NOT NULL,
  `lvl_id` int(100) NOT NULL AUTO_INCREMENT,
  `lvl_name` varchar(100) NOT NULL,
  PRIMARY KEY (`lvl_id`)
) ENGINE=InnoDB  DEFAULT CHARSET=latin1 AUTO_INCREMENT=6 ;

--
-- Dumping data for table `tbl_level`
--

INSERT INTO `tbl_level` (`SY`, `lvl_id`, `lvl_name`) VALUES
('2013-2014', 1, 'Kinder'),
('2013-2014', 2, 'Grade 1'),
('2013-2014', 3, 'Grade 2'),
('2014-2015', 4, 'Grade 1'),
('2013-2014', 5, 'Grade 3');

-- --------------------------------------------------------

--
-- Table structure for table `tbl_logs`
--

CREATE TABLE IF NOT EXISTS `tbl_logs` (
  `Username` varchar(100) NOT NULL,
  `Login` varchar(100) NOT NULL,
  `Logout` varchar(100) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tbl_logs`
--

INSERT INTO `tbl_logs` (`Username`, `Login`, `Logout`) VALUES
('admin', '1/15/2014 8:45:12 PM', '1/15/2014 8:56:17 PM'),
('admin', '1/15/2014 8:56:18 PM', '1/15/2014 8:57:25 PM'),
('admin', '1/15/2014 8:57:26 PM', '1/15/2014 9:02:42 PM'),
('admin', '1/15/2014 9:02:43 PM', '1/15/2014 9:05:21 PM'),
('admin', '1/15/2014 9:05:22 PM', '1/15/2014 9:08:05 PM'),
('admin', '1/15/2014 9:08:05 PM', '1/15/2014 9:09:03 PM'),
('M-0001Taneo', '1/15/2014 9:08:35 PM', 'None'),
('jenny', '1/15/2014 9:08:43 PM', '1/16/2014 2:53:53 AM'),
('admin', '1/15/2014 9:09:03 PM', '1/16/2014 3:02:38 AM'),
('jenny', '1/16/2014 2:53:54 AM', '1/16/2014 2:54:33 AM'),
('jenny', '1/16/2014 2:54:33 AM', '1/16/2014 3:01:38 AM'),
('jenny', '1/16/2014 3:01:38 AM', '1/16/2014 3:03:01 AM'),
('admin', '1/16/2014 3:02:38 AM', '1/16/2014 3:39:16 AM'),
('jenny', '1/16/2014 3:03:02 AM', '1/16/2014 3:10:22 AM'),
('jenny', '1/16/2014 3:10:22 AM', '1/16/2014 3:15:30 AM'),
('jenny', '1/16/2014 3:15:31 AM', '1/16/2014 3:18:41 AM'),
('jenny', '1/16/2014 3:18:42 AM', '1/16/2014 3:19:45 AM'),
('jenny', '1/16/2014 3:19:46 AM', '1/16/2014 3:26:34 AM'),
('jenny', '1/16/2014 3:26:35 AM', '1/16/2014 3:29:35 AM'),
('jenny', '1/16/2014 3:29:35 AM', '1/16/2014 3:31:34 AM'),
('jenny', '1/16/2014 3:31:34 AM', '1/16/2014 3:37:53 AM'),
('jenny', '1/16/2014 3:37:54 AM', '1/16/2014 3:38:54 AM'),
('jenny', '1/16/2014 3:38:54 AM', 'None'),
('admin', '1/16/2014 3:39:17 AM', '1/16/2014 3:51:57 AM'),
('admin', '1/16/2014 3:51:57 AM', '1/16/2014 4:04:03 AM'),
('admin', '1/16/2014 4:04:04 AM', '1/16/2014 4:10:56 AM'),
('admin', '1/16/2014 4:10:57 AM', '1/16/2014 4:12:14 AM'),
('admin', '1/16/2014 4:12:15 AM', '1/16/2014 4:12:36 AM'),
('admin', '1/16/2014 4:12:37 AM', '1/16/2014 4:13:22 AM'),
('admin', '1/16/2014 4:13:23 AM', '1/16/2014 4:13:27 AM'),
('admin', '1/16/2014 4:14:03 AM', '1/16/2014 4:15:35 AM'),
('admin', '1/16/2014 4:15:35 AM', '1/16/2014 4:21:50 AM'),
('admin', '1/16/2014 4:21:51 AM', '1/16/2014 4:26:58 AM'),
('admin', '1/16/2014 4:26:59 AM', '1/16/2014 4:32:36 AM'),
('admin', '1/16/2014 4:32:36 AM', '1/16/2014 4:48:49 AM'),
('admin', '1/16/2014 4:48:49 AM', '1/16/2014 5:08:13 AM'),
('admin', '1/16/2014 5:08:14 AM', '1/16/2014 5:17:47 AM'),
('admin', '1/16/2014 5:17:47 AM', '1/16/2014 5:32:10 AM'),
('admin', '1/16/2014 5:32:10 AM', '1/16/2014 5:45:35 AM'),
('admin', '1/16/2014 5:45:35 AM', '1/16/2014 5:53:07 AM'),
('admin', '1/16/2014 5:53:08 AM', '1/16/2014 5:53:10 AM'),
('admin', '1/16/2014 6:06:35 AM', '1/16/2014 7:20:20 AM'),
('admin', '1/16/2014 7:20:21 AM', '1/16/2014 7:36:47 AM'),
('admin', '1/16/2014 7:36:48 AM', '1/16/2014 7:48:01 AM'),
('admin', '1/16/2014 7:48:03 AM', '1/16/2014 7:53:26 AM'),
('admin', '1/16/2014 7:53:27 AM', '1/16/2014 8:06:42 AM'),
('admin', '1/16/2014 8:06:43 AM', '1/16/2014 8:08:04 AM'),
('admin', '1/16/2014 8:08:04 AM', '1/16/2014 8:18:50 AM'),
('admin', '1/16/2014 8:18:51 AM', '1/16/2014 8:30:51 AM'),
('admin', '1/16/2014 8:30:52 AM', 'None');

-- --------------------------------------------------------

--
-- Table structure for table `tbl_promotion`
--

CREATE TABLE IF NOT EXISTS `tbl_promotion` (
  `No` int(100) NOT NULL AUTO_INCREMENT,
  `Name` varchar(100) NOT NULL,
  `ID` varchar(100) NOT NULL,
  `Gender` varchar(100) NOT NULL,
  `SY` varchar(100) NOT NULL,
  `Level` varchar(100) NOT NULL,
  `Section` varchar(100) NOT NULL,
  `Address` varchar(100) NOT NULL,
  `Years_in_School` varchar(100) NOT NULL,
  `Age` varchar(100) NOT NULL,
  `Number_of_Days` varchar(100) NOT NULL,
  `Grade_Remark` varchar(100) NOT NULL,
  `Final_Rating` varchar(100) NOT NULL,
  `Action_Taken` varchar(100) NOT NULL,
  `Remark` varchar(100) NOT NULL,
  PRIMARY KEY (`No`)
) ENGINE=InnoDB  DEFAULT CHARSET=latin1 AUTO_INCREMENT=50 ;

--
-- Dumping data for table `tbl_promotion`
--

INSERT INTO `tbl_promotion` (`No`, `Name`, `ID`, `Gender`, `SY`, `Level`, `Section`, `Address`, `Years_in_School`, `Age`, `Number_of_Days`, `Grade_Remark`, `Final_Rating`, `Action_Taken`, `Remark`) VALUES
(46, 'Sample, Sample Sample', '109637333333', 'Male', '2013-2014', 'Grade 1', 'Banana', 'Sample', ' 1', ' 5.5', '0', '', 'No grade', '', ''),
(47, 'aa, a a', '109637000001', 'Female', '2013-2014', 'Grade 1', 'Apple', 'a', ' 1', ' 4.75', '0', '', 'No grade', '', ''),
(48, 'Camasoza, Charles A', '109637111111', 'Male', '2013-2014', 'Grade 1', 'Apple', 'Cavite City', ' 1', ' 6.75', ' 180', 'A', ' 90.62', '', 'A'),
(49, 'Villarde, Markprile John Bbbb', '109637222222', 'Female', '2013-2014', 'Grade 1', 'Apple', 'Cavite City', ' 1', ' 6.25', '0', 'P', ' 89.38', '', 'P');

-- --------------------------------------------------------

--
-- Table structure for table `tbl_section`
--

CREATE TABLE IF NOT EXISTS `tbl_section` (
  `SY` varchar(100) NOT NULL,
  `lvl_name` varchar(100) NOT NULL,
  `section_id` int(100) NOT NULL AUTO_INCREMENT,
  `section_name` varchar(100) NOT NULL,
  `teacher_id` varchar(100) NOT NULL,
  PRIMARY KEY (`section_id`)
) ENGINE=InnoDB  DEFAULT CHARSET=latin1 AUTO_INCREMENT=9 ;

--
-- Dumping data for table `tbl_section`
--

INSERT INTO `tbl_section` (`SY`, `lvl_name`, `section_id`, `section_name`, `teacher_id`) VALUES
('2013-2014', 'Grade 1', 1, 'Apple', 'M-0001'),
('2013-2014', 'Grade 1', 2, 'Banana', 'None'),
('2013-2014', 'Grade 1', 3, 'Cherry', 'None'),
('2013-2014', 'Grade 2', 4, 'Sampaguita', 'None'),
('2013-2014', 'Grade 2', 5, 'Rose', 'None'),
('2013-2014', 'Grade 3', 6, 'Jose Rizal', 'None'),
('2014-2015', 'Grade 1', 7, 'Apple', 'M-0001'),
('2013-2014', 'Kinder', 8, 'Dog', 'M-0001');

-- --------------------------------------------------------

--
-- Table structure for table `tbl_security`
--

CREATE TABLE IF NOT EXISTS `tbl_security` (
  `No` int(100) NOT NULL AUTO_INCREMENT,
  `Username` varchar(100) NOT NULL,
  `Pet` varchar(100) NOT NULL,
  `Place` varchar(100) NOT NULL,
  `Author` varchar(100) NOT NULL,
  PRIMARY KEY (`No`)
) ENGINE=InnoDB  DEFAULT CHARSET=latin1 AUTO_INCREMENT=3 ;

--
-- Dumping data for table `tbl_security`
--

INSERT INTO `tbl_security` (`No`, `Username`, `Pet`, `Place`, `Author`) VALUES
(1, 'admin', 'aso', 'bahay', 'ako'),
(2, 'M-0001Taneo', 'a', 'a', 'a');

-- --------------------------------------------------------

--
-- Table structure for table `tbl_student`
--

CREATE TABLE IF NOT EXISTS `tbl_student` (
  `No` int(100) NOT NULL AUTO_INCREMENT,
  `student_id` varchar(200) NOT NULL,
  `last_name` varchar(200) NOT NULL,
  `first_name` varchar(200) NOT NULL,
  `middle_name` varchar(200) NOT NULL,
  `gender` varchar(200) NOT NULL,
  `bday` varchar(200) NOT NULL,
  `birthplace` varchar(100) NOT NULL,
  `contact_no` varchar(200) NOT NULL,
  `address` varchar(200) NOT NULL,
  `guardian` varchar(200) NOT NULL,
  `guardian_no` varchar(200) NOT NULL,
  `occupation` varchar(100) NOT NULL,
  PRIMARY KEY (`No`)
) ENGINE=InnoDB  DEFAULT CHARSET=latin1 AUTO_INCREMENT=6 ;

--
-- Dumping data for table `tbl_student`
--

INSERT INTO `tbl_student` (`No`, `student_id`, `last_name`, `first_name`, `middle_name`, `gender`, `bday`, `birthplace`, `contact_no`, `address`, `guardian`, `guardian_no`, `occupation`) VALUES
(1, '109637111111', 'Camasoza', 'Charles', 'A', 'Male', '2007-03-04', 'Cavite City', '1', 'Cavite City', 'A Camasoza', '2333', 'Housewife'),
(2, '109637222222', 'Villarde', 'Markprile John', 'Bbbb', 'Female', '2007-08-04', '', '1', 'Cavite City', 'A Villarde', '2', ''),
(4, '109637333333', 'Sample', 'Sample', 'Sample', 'Male', '2008-07-04', '', 'Sample', 'Sample', 'Sample', 'Sample', ''),
(5, '109637000001', 'To', 'Bata', 'a', 'Female', '2009-02-12', '', '1', 'a', 'a', '1', '');

-- --------------------------------------------------------

--
-- Table structure for table `tbl_student_level`
--

CREATE TABLE IF NOT EXISTS `tbl_student_level` (
  `No` int(100) NOT NULL AUTO_INCREMENT,
  `ID` varchar(200) NOT NULL,
  `SY` varchar(200) NOT NULL,
  `lvl_name` varchar(200) NOT NULL,
  `section_name` varchar(200) NOT NULL,
  `Status` varchar(100) NOT NULL,
  PRIMARY KEY (`No`)
) ENGINE=InnoDB  DEFAULT CHARSET=latin1 AUTO_INCREMENT=6 ;

--
-- Dumping data for table `tbl_student_level`
--

INSERT INTO `tbl_student_level` (`No`, `ID`, `SY`, `lvl_name`, `section_name`, `Status`) VALUES
(1, '109637111111', '2013-2014', 'Grade 1', 'Apple', 'ENROLLED'),
(2, '109637222222', '2013-2014', 'Grade 1', 'Apple', 'ENROLLED'),
(4, '109637333333', '2013-2014', 'Grade 1', 'Banana', 'ENROLLED'),
(5, '109637000001', '2013-2014', 'Grade 1', 'Apple', 'ENROLLED');

-- --------------------------------------------------------

--
-- Table structure for table `tbl_subject`
--

CREATE TABLE IF NOT EXISTS `tbl_subject` (
  `No` int(100) NOT NULL AUTO_INCREMENT,
  `SY` varchar(100) NOT NULL,
  `lvl_name` varchar(100) NOT NULL,
  `subject_code` varchar(100) NOT NULL,
  `subject_name` varchar(100) NOT NULL,
  PRIMARY KEY (`No`)
) ENGINE=InnoDB  DEFAULT CHARSET=latin1 AUTO_INCREMENT=7 ;

--
-- Dumping data for table `tbl_subject`
--

INSERT INTO `tbl_subject` (`No`, `SY`, `lvl_name`, `subject_code`, `subject_name`) VALUES
(1, '2013-2014', 'Grade 1', 'Eng', 'English'),
(2, '2013-2014', 'Grade 2', 'Fil', 'Filipino'),
(4, '2013-2014', 'Grade 1', 'Fil', 'Filipino'),
(5, '2013-2014', 'Grade 1', 'Mat', 'Mathematics'),
(6, '2013-2014', 'Kinder', 'Math', 'Mathematics');

-- --------------------------------------------------------

--
-- Table structure for table `tbl_subjectset`
--

CREATE TABLE IF NOT EXISTS `tbl_subjectset` (
  `No` int(11) NOT NULL AUTO_INCREMENT,
  `SY` varchar(100) NOT NULL,
  `lvl_name` varchar(100) NOT NULL,
  `section_name` varchar(100) NOT NULL,
  `subject_code` varchar(100) NOT NULL,
  `subject_name` varchar(100) NOT NULL,
  `teacher_id` varchar(100) NOT NULL,
  PRIMARY KEY (`No`)
) ENGINE=InnoDB  DEFAULT CHARSET=latin1 AUTO_INCREMENT=4 ;

--
-- Dumping data for table `tbl_subjectset`
--

INSERT INTO `tbl_subjectset` (`No`, `SY`, `lvl_name`, `section_name`, `subject_code`, `subject_name`, `teacher_id`) VALUES
(1, '2013-2014', 'Grade 1', 'Apple', 'Eng', 'English', 'M-0001'),
(2, '2013-2014', 'Grade 1', 'Apple', 'Fil', 'Filipino', 'M-0001'),
(3, '2013-2014', 'Grade 1', 'Apple', 'Mat', 'Mathematics', 'M-0002');

-- --------------------------------------------------------

--
-- Table structure for table `tbl_sy`
--

CREATE TABLE IF NOT EXISTS `tbl_sy` (
  `SY` varchar(100) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tbl_sy`
--

INSERT INTO `tbl_sy` (`SY`) VALUES
('2013'),
('2014'),
('2012'),
('2010');

-- --------------------------------------------------------

--
-- Table structure for table `tbl_teacher`
--

CREATE TABLE IF NOT EXISTS `tbl_teacher` (
  `No` int(11) NOT NULL AUTO_INCREMENT,
  `teacher_id` varchar(100) NOT NULL,
  `first_name` varchar(200) NOT NULL,
  `last_name` varchar(200) NOT NULL,
  `middle_name` varchar(100) NOT NULL,
  `gender` varchar(200) NOT NULL,
  `bday` varchar(200) NOT NULL,
  `contact_no` varchar(200) NOT NULL,
  `course` varchar(200) NOT NULL,
  `school` varchar(200) NOT NULL,
  `a_from` varchar(100) NOT NULL,
  `a_to` varchar(100) NOT NULL,
  `address` varchar(100) NOT NULL,
  `status` varchar(200) NOT NULL,
  PRIMARY KEY (`No`)
) ENGINE=InnoDB  DEFAULT CHARSET=latin1 AUTO_INCREMENT=4 ;

--
-- Dumping data for table `tbl_teacher`
--

INSERT INTO `tbl_teacher` (`No`, `teacher_id`, `first_name`, `last_name`, `middle_name`, `gender`, `bday`, `contact_no`, `course`, `school`, `a_from`, `a_to`, `address`, `status`) VALUES
(1, 'M-0001', 'Jenny', 'Taneo', 'Aaaa', 'Female', '1996-12-04', '1111', 'Bachelor of Science in Information Technology', 'Cavite State University - Cavite City Campus', '2009', '2013', 'Cavite City', 'On-Duty'),
(2, 'M-0002', 'Danie Anne', 'Pamatian', 'Bbb', 'Female', '1994-05-04', '2222', 'Bachelor of Scuence in Computer Science', 'AMA', '2000', '2004', 'Cavite City', 'On-Duty'),
(3, 'M-0003', 'Sample', 'Sample', 'Aaaabba', 'Female', '2014-01-12', '111111111111', '', '', '', '', '', 'On-Duty');

-- --------------------------------------------------------

--
-- Table structure for table `tbl_user`
--

CREATE TABLE IF NOT EXISTS `tbl_user` (
  `No` int(100) NOT NULL AUTO_INCREMENT,
  `ID` varchar(100) NOT NULL,
  `Usertype` varchar(100) NOT NULL,
  `Username` varchar(100) NOT NULL,
  `Password` varchar(100) NOT NULL,
  `Lastname` varchar(200) NOT NULL,
  `Firstname` varchar(200) NOT NULL,
  `Middlename` varchar(200) NOT NULL,
  PRIMARY KEY (`No`)
) ENGINE=InnoDB  DEFAULT CHARSET=latin1 AUTO_INCREMENT=6 ;

--
-- Dumping data for table `tbl_user`
--

INSERT INTO `tbl_user` (`No`, `ID`, `Usertype`, `Username`, `Password`, `Lastname`, `Firstname`, `Middlename`) VALUES
(1, '1', 'Administrator', 'admin', 'admin', 'admin', 'admin', 'admin'),
(2, 'M-0001', 'Teacher', 'jenny', 'jenny', 'Taneo', 'Jenny', 'Aaaa'),
(3, 'M-0002', 'Teacher', 'M-0002Pamatian', 'Pamatian', 'Pamatian', 'Danie Anne', 'Bbb'),
(5, 'M-0003', 'Teacher', 'M-0003Sample', 'Sample', '', '', '');

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
