﻿Public NotInheritable Class SplashScreen1

    'TODO: This form can easily be set as the splash screen for the application by going to the "Application" tab
    '  of the Project Designer ("Properties" under the "Project" menu).


    Private Sub SplashScreen1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Set up the dialog text at runtime according to the application's assembly information.  

        'TODO: Customize the application's assembly information in the "Application" pane of the project 
        '  properties dialog (under the "Project" menu).

        'Application title

        'If the application title is missing, use the application name, without the extension



        'Format the version information using the text set into the Version control at design time as the
        '  formatting string.  This allows for effective localization if desired.
        '  Build and revision information could be included by using the following code and changing the 
        '  Version control's designtime text to "Version {0}.{1:00}.{2}.{3}" or something similar.  See
        '  String.Format() in Help for more information.
        '
        '    Version.Text = System.String.Format(Version.Text, My.Application.Info.Version.Major, My.Application.Info.Version.Minor, My.Application.Info.Version.Build, My.Application.Info.Version.Revision)



        'Copyright info


    End Sub

    Private Sub ApplicationTitle_Click(sender As Object, e As EventArgs)

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub SplashScreen1_Validated(sender As Object, e As EventArgs) Handles Me.Validated

    End Sub

    Private Sub MainLayoutPanel_Paint(sender As Object, e As PaintEventArgs)

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs)

    End Sub
End Class
