CREATE DATABASE  IF NOT EXISTS `db_form137` /*!40100 DEFAULT CHARACTER SET latin1 */;
USE `db_form137`;
-- MySQL dump 10.13  Distrib 5.6.13, for Win32 (x86)
--
-- Host: 127.0.0.1    Database: db_form137
-- ------------------------------------------------------
-- Server version	5.6.12-log

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;
/*!40103 SET @OLD_TIME_ZONE=@@TIME_ZONE */;
/*!40103 SET TIME_ZONE='+00:00' */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;

--
-- Table structure for table `tbl_promotion`
--

DROP TABLE IF EXISTS `tbl_promotion`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_promotion` (
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
) ENGINE=InnoDB AUTO_INCREMENT=73 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_promotion`
--

LOCK TABLES `tbl_promotion` WRITE;
/*!40000 ALTER TABLE `tbl_promotion` DISABLE KEYS */;
INSERT INTO `tbl_promotion` VALUES (59,'Sample, Sample Sample','109637333333','Male','2013-2014','Grade 1','Banana','Sample',' 1',' 5.5','0','','No grade','',''),(63,'Magsaysay, Rexa Acel ','109637000006','Female','','Grade 1','Cherry','Cavite City',' 1',' 10.75','0','P',' 89.5','','P'),(70,'Camasoza, Charles A','109637111111','Male','','Grade 1','Apple','Cavite City',' 1',' 6.75',' 180','A',' 90.62','','A'),(71,'To, Bata a','109637000001','Female','','Grade 1','Apple','a',' 1',' 4.75','0','A','No grade','','A'),(72,'Villarde, Markprile John Bbbb','109637222222','Female','','Grade 1','Apple','Cavite City',' 1',' 6.25','0','P',' 89.38','','P');
/*!40000 ALTER TABLE `tbl_promotion` ENABLE KEYS */;
UNLOCK TABLES;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2014-01-28 18:48:20
