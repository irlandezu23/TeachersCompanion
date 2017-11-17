
-- --------------------------------------------------
-- Entity Designer DDL Script for SQL Server 2005, 2008, 2012 and Azure
-- --------------------------------------------------
-- Date Created: 10/13/2016 15:57:33
-- Generated from EDMX file: C:\Users\Catalin\Documents\Visual Studio 2015\Projects\Teacherscompanion\Teacherscompanion\TCdbmodel.edmx
-- --------------------------------------------------

SET QUOTED_IDENTIFIER OFF;
GO
USE [C:\Users\Catalin\Documents\Visual Studio 2015\Projects\Teacherscompanion\mydatabase.mdf];
GO
IF SCHEMA_ID(N'dbo') IS NULL EXECUTE(N'CREATE SCHEMA [dbo]');
GO

-- --------------------------------------------------
-- Dropping existing FOREIGN KEY constraints
-- --------------------------------------------------

IF OBJECT_ID(N'[dbo].[FK_ClassStudents]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[Students] DROP CONSTRAINT [FK_ClassStudents];
GO

-- --------------------------------------------------
-- Dropping existing tables
-- --------------------------------------------------

IF OBJECT_ID(N'[dbo].[Classes]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Classes];
GO
IF OBJECT_ID(N'[dbo].[Students]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Students];
GO

-- --------------------------------------------------
-- Creating all tables
-- --------------------------------------------------

-- Creating table 'Classes'
CREATE TABLE [dbo].[Classes] (
    [Id] int  NOT NULL
);
GO

-- Creating table 'Students'
CREATE TABLE [dbo].[Students] (
    [RFid] int  NOT NULL,
    [Firstname] nvarchar(max)  NOT NULL,
    [Surname] nvarchar(max)  NOT NULL,
    [ClassId] int  NOT NULL
);
GO

-- Creating table 'Sessions'
CREATE TABLE [dbo].[Sessions] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [Date] datetime  NOT NULL,
    [ClassesId] int  NOT NULL
);
GO

-- Creating table 'Logs'
CREATE TABLE [dbo].[Logs] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [StudentsRFid] int  NOT NULL,
    [SessionsId] int  NOT NULL
);
GO

-- --------------------------------------------------
-- Creating all PRIMARY KEY constraints
-- --------------------------------------------------

-- Creating primary key on [Id] in table 'Classes'
ALTER TABLE [dbo].[Classes]
ADD CONSTRAINT [PK_Classes]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [RFid] in table 'Students'
ALTER TABLE [dbo].[Students]
ADD CONSTRAINT [PK_Students]
    PRIMARY KEY CLUSTERED ([RFid] ASC);
GO

-- Creating primary key on [Id] in table 'Sessions'
ALTER TABLE [dbo].[Sessions]
ADD CONSTRAINT [PK_Sessions]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Logs'
ALTER TABLE [dbo].[Logs]
ADD CONSTRAINT [PK_Logs]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- --------------------------------------------------
-- Creating all FOREIGN KEY constraints
-- --------------------------------------------------

-- Creating foreign key on [ClassId] in table 'Students'
ALTER TABLE [dbo].[Students]
ADD CONSTRAINT [FK_ClassStudents]
    FOREIGN KEY ([ClassId])
    REFERENCES [dbo].[Classes]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_ClassStudents'
CREATE INDEX [IX_FK_ClassStudents]
ON [dbo].[Students]
    ([ClassId]);
GO

-- Creating foreign key on [ClassesId] in table 'Sessions'
ALTER TABLE [dbo].[Sessions]
ADD CONSTRAINT [FK_ClassesSessions]
    FOREIGN KEY ([ClassesId])
    REFERENCES [dbo].[Classes]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_ClassesSessions'
CREATE INDEX [IX_FK_ClassesSessions]
ON [dbo].[Sessions]
    ([ClassesId]);
GO

-- Creating foreign key on [StudentsRFid] in table 'Logs'
ALTER TABLE [dbo].[Logs]
ADD CONSTRAINT [FK_StudentsLogs]
    FOREIGN KEY ([StudentsRFid])
    REFERENCES [dbo].[Students]
        ([RFid])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_StudentsLogs'
CREATE INDEX [IX_FK_StudentsLogs]
ON [dbo].[Logs]
    ([StudentsRFid]);
GO

-- Creating foreign key on [SessionsId] in table 'Logs'
ALTER TABLE [dbo].[Logs]
ADD CONSTRAINT [FK_SessionsLogs]
    FOREIGN KEY ([SessionsId])
    REFERENCES [dbo].[Sessions]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_SessionsLogs'
CREATE INDEX [IX_FK_SessionsLogs]
ON [dbo].[Logs]
    ([SessionsId]);
GO

-- --------------------------------------------------
-- Script has ended
-- --------------------------------------------------