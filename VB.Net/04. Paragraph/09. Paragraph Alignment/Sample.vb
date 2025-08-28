Imports System
Imports SautinSoft.Document
Imports System.Text
Imports SautinSoft.Document.Drawing
Imports System.Reflection.Metadata

Namespace Example
    Class Program
        Public Shared Sub Main(ByVal args As String())
            ParagraphAlignment()
        End Sub

        Private Shared Sub ParagraphAlignment()
            Dim resultPath As String = "..\..\..\Result.docx"
            Dim pictPath As String = "..\..\..\logo.png"
            Dim SpecialIntendt As Integer = 35
            Dim dc As DocumentCore = New DocumentCore()
            Dim db As DocumentBuilder = New DocumentBuilder(dc)
            Dim paragraphStyle As ParagraphStyle = New ParagraphStyle("ParagraphStyle1")
            paragraphStyle.CharacterFormat.FontName = "Arial Narrow"
            paragraphStyle.CharacterFormat.Size = 10
            paragraphStyle.ParagraphFormat.Alignment = SautinSoft.Document.HorizontalAlignment.Left
            paragraphStyle.ParagraphFormat.SpecialIndentation = SpecialIntendt
            Dim paragraphStyleJust As ParagraphStyle = New ParagraphStyle("ParagraphStyle2")
            paragraphStyleJust.CharacterFormat.FontName = "Arial Narrow"
            paragraphStyleJust.CharacterFormat.Size = 10
            paragraphStyleJust.ParagraphFormat.Alignment = SautinSoft.Document.HorizontalAlignment.Justify
            Dim paragraphStyleRight As ParagraphStyle = New ParagraphStyle("ParagraphStyle3")
            paragraphStyleRight.CharacterFormat.FontName = "Arial Narrow"
            paragraphStyleRight.CharacterFormat.Size = 10
            paragraphStyleRight.ParagraphFormat.Alignment = SautinSoft.Document.HorizontalAlignment.Right
            Dim characterStyle As CharacterStyle = New CharacterStyle("CharacterStyle1")
            characterStyle.CharacterFormat.FontName = "Arial Narrow"
            characterStyle.CharacterFormat.UnderlineStyle = UnderlineType.Single
            characterStyle.CharacterFormat.Size = 10
            dc.Styles.Add(paragraphStyle)
            dc.Styles.Add(paragraphStyleJust)
            dc.Styles.Add(paragraphStyleRight)
            dc.Styles.Add(characterStyle)
            Dim pict2 As Picture = db.InsertImage(pictPath, New HorizontalPosition(-7.5, LengthUnit.Centimeter, HorizontalPositionAnchor.RightMargin), New VerticalPosition(1, LengthUnit.Centimeter, VerticalPositionAnchor.TopMargin), WrappingStyle.InFrontOfText, New Size(65, 15, LengthUnit.Millimeter))
            Dim p As Paragraph = New Paragraph(dc, "COMMUNITY ASSOCIATION AUTHORITY TO REPRESENT")
            p.ParagraphFormat.Alignment = SautinSoft.Document.HorizontalAlignment.Center
            dc.Sections(0).Blocks.Add(p)
            Dim p2 As Paragraph = New Paragraph(dc, "AND RETAINER AGREEMENT")
            p2.ParagraphFormat.Alignment = SautinSoft.Document.HorizontalAlignment.Center
            dc.Sections(0).Blocks.Add(p2)
            Dim p3 As Paragraph = New Paragraph(dc, "This Agreement is effective from January 1, 2025 to December 31, 2029.") With {
                .ParagraphFormat = New ParagraphFormat With {
                    .Style = paragraphStyle
                }
            }
            dc.Sections(0).Blocks.Add(p3)
            Dim p4 As Paragraph = New Paragraph(dc, "This Agreement can be renewed for additional one year periods, AT THE ELECTION OF THE ASSOCIATION, by paying the Annual Retainer Fee for the next calendar year period.  The billing rate and the terms of this Agreement  effective for any renewal term remain the same, unless a change is announced by the Firm in writing.  This Agreement may be terminated at any time by either party by giving written notice to the other party.") With {
                .ParagraphFormat = New ParagraphFormat With {
                    .Style = paragraphStyle
                }
            }
            dc.Sections(0).Blocks.Add(p4)
            Dim p5 As Paragraph = New Paragraph(dc, """Company__Also_Known_As_""   (""Association""),") With {
                .ParagraphFormat = New ParagraphFormat With {
                    .Alignment = SautinSoft.Document.HorizontalAlignment.Center
                }
            }
            dc.Sections(0).Blocks.Add(p5)
            Dim p6 As Paragraph = New Paragraph(dc, "by and through its Board of Directors, retains the law firm of TruckCom & Sautin, P.A. (hereinafter referred to as ""Firm"" or ""Becker""), to represent it as legal counsel in the matters described in this Agreement and its Exhibits. This retainer and representation is solely and exclusively for the benefit of the Association, as a corporate entity, and not for any other party or third parties.  There are no intended third party beneficiaries, including, but not limited to members of the Association; residents living in the community operated by the Association or guests of such residents; officers, directors, employees, or agents of the Association.") With {
                .ParagraphFormat = New ParagraphFormat With {
                    .Style = paragraphStyleJust
                }
            }
            dc.Sections(0).Blocks.Add(p6)
            Dim p7 As Paragraph = New Paragraph(dc, "Paying the Annual Retainer Fee sum of Three Hundred Dollars ($300.00)  entitles the Association to the services listed in Exhibit ""A"" at no extra charge.") With {
                .ParagraphFormat = New ParagraphFormat With {
                    .Style = paragraphStyleJust
                }
            }
            dc.Sections(0).Blocks.Add(p7)
            Dim p8 As Paragraph = New Paragraph(dc, "The Firm will provide general legal services upon the request of the Association or as authorized by this Agreement concerning the day-to-day operation of the Association, including certain litigation, arbitration and mediation matters, on the reduced hourly fees stated in Exhibit ""B"" to this Agreement, subject to the terms and conditions in Exhibit ""A"".") With {
                .ParagraphFormat = New ParagraphFormat With {
                    .Style = paragraphStyleJust
                }
            }
            dc.Sections(0).Blocks.Add(p8)
            Dim p9 As Paragraph = New Paragraph(dc, "The undersigned officer or agent of the Association represents and certifies that he(she) is authorized by the Board of Directors to execute this Agreement and agrees to the terms contained in the attached Exhibits on behalf of the Association and its membership.") With {
                .ParagraphFormat = New ParagraphFormat With {
                    .Style = paragraphStyleJust
                }
            }
            dc.Sections(0).Blocks.Add(p9)
            Dim p10 As Paragraph = New Paragraph(dc, """ASSOCIATION""") With {
                .ParagraphFormat = New ParagraphFormat With {
                    .Style = paragraphStyleRight
                }
            }
            dc.Sections(0).Blocks.Add(p10)
            Dim p11 As Paragraph = New Paragraph(dc, "By:_________________________") With {
                .ParagraphFormat = New ParagraphFormat With {
                    .Style = paragraphStyleRight
                }
            }
            dc.Sections(0).Blocks.Add(p11)
            Dim p12 As Paragraph = New Paragraph(dc, "Print Nane:_________________________") With {
                .ParagraphFormat = New ParagraphFormat With {
                    .Style = paragraphStyleRight
                }
            }
            dc.Sections(0).Blocks.Add(p12)
            Dim p13 As Paragraph = New Paragraph(dc, "Print Title:_________________________") With {
                .ParagraphFormat = New ParagraphFormat With {
                    .Style = paragraphStyleRight
                }
            }
            dc.Sections(0).Blocks.Add(p13)
            Dim p114 As Paragraph = New Paragraph(dc, "Date:_________________________") With {
                .ParagraphFormat = New ParagraphFormat With {
                    .Style = paragraphStyleRight
                }
            }
            dc.Sections(0).Blocks.Add(p114)
            Dim p115 As Paragraph = New Paragraph(dc)
            p115.ParagraphFormat.Tabs.Add(New TabStop(350, TabStopAlignment.Left))
            p115.Inlines.Add(New SpecialCharacter(dc, SpecialCharacterType.Tab))
            p115.Inlines.Add(New Run(dc, "By______________"))
            dc.Sections(0).Blocks.Add(p115)
            Dim p116 As Paragraph = New Paragraph(dc)
            p116.ParagraphFormat.Tabs.Add(New TabStop(350, TabStopAlignment.Left))
            p116.Inlines.Add(New SpecialCharacter(dc, SpecialCharacterType.Tab))
            p116.Inlines.Add(New Run(dc, "Print Name__________"))
            dc.Sections(0).Blocks.Add(p116)
            Dim p117 As Paragraph = New Paragraph(dc)
            p117.ParagraphFormat.Tabs.Add(New TabStop(350, TabStopAlignment.Left))
            p117.Inlines.Add(New SpecialCharacter(dc, SpecialCharacterType.Tab))
            p117.Inlines.Add(New Run(dc, "Print Title____________"))
            dc.Sections(0).Blocks.Add(p117)
            Dim p118 As Paragraph = New Paragraph(dc)
            p118.ParagraphFormat.Tabs.Add(New TabStop(350, TabStopAlignment.Left))
            p118.Inlines.Add(New SpecialCharacter(dc, SpecialCharacterType.Tab))
            p118.Inlines.Add(New Run(dc, "Date___________"))
            dc.Sections(0).Blocks.Add(p118)
            Dim p119 As Paragraph = New Paragraph(dc)
            p119.ParagraphFormat.Tabs.Add(New TabStop(350, TabStopAlignment.Center))
            p119.Inlines.Add(New SpecialCharacter(dc, SpecialCharacterType.Tab))
            p119.Inlines.Add(New Run(dc, "By______________"))
            dc.Sections(0).Blocks.Add(p119)
            Dim p120 As Paragraph = New Paragraph(dc)
            p120.ParagraphFormat.Tabs.Add(New TabStop(350, TabStopAlignment.Center))
            p120.Inlines.Add(New SpecialCharacter(dc, SpecialCharacterType.Tab))
            p120.Inlines.Add(New Run(dc, "Print Name__________"))
            dc.Sections(0).Blocks.Add(p120)
            Dim p121 As Paragraph = New Paragraph(dc)
            p121.ParagraphFormat.Tabs.Add(New TabStop(350, TabStopAlignment.Center))
            p121.Inlines.Add(New SpecialCharacter(dc, SpecialCharacterType.Tab))
            p121.Inlines.Add(New Run(dc, "Print Title____________"))
            dc.Sections(0).Blocks.Add(p121)
            Dim p122 As Paragraph = New Paragraph(dc)
            p122.ParagraphFormat.Tabs.Add(New TabStop(350, TabStopAlignment.Center))
            p122.Inlines.Add(New SpecialCharacter(dc, SpecialCharacterType.Tab))
            p122.Inlines.Add(New Run(dc, "Date___________"))
            dc.Sections(0).Blocks.Add(p122)
            dc.Save(resultPath, New DocxSaveOptions())
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(resultPath) With {
                .UseShellExecute = True
            })
        End Sub
    End Class
End Namespace
