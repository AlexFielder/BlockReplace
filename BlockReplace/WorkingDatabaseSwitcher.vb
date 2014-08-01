Imports Autodesk.AutoCAD.DatabaseServices
''' <summary>
''' change the database of drawings, it is important to control which database is current.
''' Class <c> WorkingDatabaseSwitcher </ c>
''' assumes control over that was exactly the current database that is needed.
''' </ summary>
''' <example>
''' Example of the class
''' <code>
''' / / db - Database object
''' using (WorkingDatabaseSwitcher hlp = New WorkingDatabaseSwitcher (db)) {
''' / / our code here </ code>
'''} </ example>
Class WorkingDatabaseSwitcher : Implements IDisposable

    Private prevDb As Database
    ''' <summary>
    ''' database, in the context of work that must be done. This database is current at the time.
    ''' At the end of the current database will be the one that was before it.
    ''' </ summary>
    ''' <param name="db"> database that must be set current </ param>
    Sub New(db As Database)
        ' TODO: Complete member initialization 
        prevDb = HostApplicationServices.WorkingDatabase
        HostApplicationServices.WorkingDatabase = db
    End Sub
    ''' <summary>
    ''' Resets the HostApplicationServices.WorkingDatabase to the previous value
    ''' </summary>
    ''' <remarks></remarks>
    Sub Dispose() Implements IDisposable.Dispose
        HostApplicationServices.WorkingDatabase = prevDb
    End Sub

End Class

