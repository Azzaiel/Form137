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
-- Table structure for table `tbl_user`
--

DROP TABLE IF EXISTS `tbl_user`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_user` (
  `No` int(100) NOT NULL AUTO_INCREMENT,
  `ID` varchar(100) NOT NULL,
  `Usertype` varchar(100) NOT NULL,
  `Username` varchar(100) NOT NULL,
  `Password` varchar(100) NOT NULL,
  `Lastname` varchar(200) NOT NULL,
  `Firstname` varchar(200) NOT NULL,
  `Middlename` varchar(200) NOT NULL,
  PRIMARY KEY (`No`)
) ENGINE=InnoDB AUTO_INCREMENT=14 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_user`
--

LOCK TABLES `tbl_user` WRITE;
/*!40000 ALTER TABLE `tbl_user` DISABLE KEYS */;
INSERT INTO `tbl_user` VALUES (1,'1','Administrator','admin','admin','admin','admin','admin'),(2,'M-0001','Teacher','jenny','Taneo','Taneo','Jenny','Aaaa'),(3,'M-0002','Teacher','M-0002Pamatian','Pamatian','Pamatian','Danie Anne','Bbb'),(5,'M-0003','Teacher','M-0003Sample','Sample','','',''),(6,'M-0004','Teacher','M-0004Mascardo','Mascardo','','',''),(7,'M-0005','Teacher','M-0005asdsdaf','asdsdaf','','',''),(8,'M-0006','Teacher','M-0006dddd','dddd','','',''),(9,'M-0007','Teacher','M-0007adasd','adasd','','',''),(10,'M-0008','Teacher','M-0008www','www','','',''),(11,'M-0009','Teacher','M-0009sss','sss','','',''),(12,'M-0010','Teacher','M-0010adasd','adasd','','',''),(13,'M-0011','Teacher','M-0011chua','chua','','','');
/*!40000 ALTER TABLE `tbl_user` ENABLE KEYS */;
UNLOCK TABLES;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2014-01-28 18:48:21
