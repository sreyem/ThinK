
Imports essentials

Module Module1

    Sub Main()

        Dim frm As New frmPGrid(
            class2Show:=test,
            classType:=test.GetType)

        Try
            With frm

                .Width = 900
                .Height = 900
                .Text = "Soil"
                .ShowDialog()

            End With

            test.CopyPropertiesByName(src:=frm.class2Show)

        Catch ex As Exception

        End Try

    End Sub

    Public Property test As New soilTexture

End Module
