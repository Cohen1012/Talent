USE [TalentDatabase]
GO
/****** Object:  View [dbo].[FilterTable]    Script Date: 2018/1/24 下午 06:52:33 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[FilterTable]
AS
SELECT          TOP (100) PERCENT dbo.Code.Code_Id, dbo.Contact_Info.Name, dbo.Contact_Info.Sex, dbo.Contact_Info.Mail, 
                            dbo.Contact_Info.CellPhone, dbo.Contact_Info.UpdateTime, dbo.Contact_Info.Cooperation_Mode, 
                            dbo.Contact_Info.Status, dbo.Contact_Info.Place, dbo.Contact_Info.Skill, dbo.Contact_Situation.Contact_Status, 
                            dbo.Contact_Situation.Remarks, dbo.Contact_Situation.Contact_Date, dbo.Interview_Info.Vacancies, 
                            dbo.Interview_Info.Interview_Date, dbo.Interview_Info.Name AS Interview_Info_Name, 
                            dbo.Interview_Info.Sex AS Interview_Info_Sex, dbo.Interview_Info.Birthday, dbo.Interview_Info.Married, 
                            dbo.Interview_Info.Mail AS Interview_Info_Mail, dbo.Interview_Info.Adress, 
                            dbo.Interview_Info.CellPhone AS Interview_Info_CellPhone, dbo.Interview_Info.Image, 
                            dbo.Interview_Info.Expertise_Language, dbo.Interview_Info.Expertise_Tools, dbo.Interview_Info.Expertise_Devops, 
                            dbo.Interview_Info.Expertise_OS, dbo.Interview_Info.Expertise_BigData, dbo.Interview_Info.Expertise_DataBase, 
                            dbo.Interview_Info.Expertise_Certification, dbo.Interview_Info.IsStudy, dbo.Interview_Info.IsService, 
                            dbo.Interview_Info.Relatives_Relationship, dbo.Interview_Info.Relatives_Name, dbo.Interview_Info.Care_Work, 
                            dbo.Interview_Info.Hope_Salary, dbo.Interview_Info.When_Report, dbo.Interview_Info.Advantage, 
                            dbo.Interview_Info.Disadvantages, dbo.Interview_Info.Hobby, dbo.Interview_Info.Attract_Reason, 
                            dbo.Interview_Info.Future_Goal, dbo.Interview_Info.Hope_Supervisor, dbo.Interview_Info.Hope_Promise, 
                            dbo.Interview_Info.Introduction, dbo.Interview_Info.Appointment, dbo.Interview_Info.Results_Remark, 
                            dbo.Interview_Info.During_Service, dbo.Interview_Info.Exemption_Reason, 
                            dbo.Interview_Info.Urgent_Contact_Person, dbo.Interview_Info.Urgent_Relationship, 
                            dbo.Interview_Info.Urgent_CellPhone, dbo.Interview_Info.Education, dbo.Interview_Info.Language, 
                            dbo.Interview_Info.Work_Experience, dbo.Interview_Comments.Interviewer, dbo.Interview_Comments.Result, 
                            dbo.Project_Experience.Company, dbo.Project_Experience.Project_Name, dbo.Project_Experience.OS, 
                            dbo.Project_Experience.[Database], dbo.Project_Experience.Position, dbo.Project_Experience.Language AS Expr10, 
                            dbo.Project_Experience.Tools, dbo.Project_Experience.Description, dbo.Project_Experience.Start_End_Date, 
                            dbo.Code.Id, dbo.Interview_Info.Interview_Id, dbo.Contact_Info.Contact_Id
FROM              dbo.Contact_Info LEFT OUTER JOIN
                            dbo.Code ON dbo.Contact_Info.Contact_Id = dbo.Code.Contact_Id LEFT OUTER JOIN
                            dbo.Contact_Situation ON dbo.Contact_Situation.Contact_Id = dbo.Code.Contact_Id LEFT OUTER JOIN
                            dbo.Interview_Info ON dbo.Interview_Info.Contact_Id = dbo.Code.Contact_Id LEFT OUTER JOIN
                            dbo.Project_Experience ON dbo.Interview_Info.Interview_Id = dbo.Project_Experience.Interview_Id LEFT OUTER JOIN
                            dbo.Interview_Comments ON dbo.Project_Experience.Interview_Id = dbo.Interview_Comments.Interview_Id
ORDER BY   dbo.Contact_Info.Contact_Id, dbo.Contact_Situation.Contact_Date DESC, dbo.Interview_Info.Interview_Date DESC
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[53] 4[1] 2[33] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "Code"
            Begin Extent = 
               Top = 6
               Left = 10
               Bottom = 119
               Right = 175
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "Contact_Info"
            Begin Extent = 
               Top = 0
               Left = 206
               Bottom = 233
               Right = 407
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "Contact_Situation"
            Begin Extent = 
               Top = 129
               Left = 5
               Bottom = 278
               Right = 174
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "Interview_Comments"
            Begin Extent = 
               Top = 0
               Left = 704
               Bottom = 113
               Right = 869
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "Interview_Info"
            Begin Extent = 
               Top = 1
               Left = 434
               Bottom = 166
               Right = 654
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "Project_Experience"
            Begin Extent = 
               Top = 180
               Left = 685
               Bottom = 321
               Right = 859
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 76
         Width = 284
         Width ' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'FilterTable'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane2', @value=N'= 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1470
         Alias = 900
         Table = 1710
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'FilterTable'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=2 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'FilterTable'
GO
