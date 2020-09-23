VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPluginInterface"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'This class wil have be be the stnadard object you use to get the plugin going.
' You can edd extra functions the perform other things, like
' KillPlugin, HidePlugin etc etc. The basic rule will just be that these functions
' must be known to the host app. How else will the host app know how to
' get the plugin going ?

' Note that we pass the host app's form into this object. Once we
' got the host app's form, we have access to the host app's form
' and we cann access and manipulate all public objects on that
' form. An Example of why we do this is : Think of a database application.
' The DB application uses ADO. You want to write a pluin that would
' calculate the total number of characters in a record set. So,
' the ADO RECORDSET (Resultset) object must be public on the host
' form. By passing the form to the plugin, the plugin can get to the ADO
' object and do the calculations. Thus, the plugin is specific for
' a certain form, becuase it needs to know what goes on on that form.


' Got that ?



Public Function Run(FormX As Object) As String
'This is the entry point to run the plugin. The Host app will have
' to call this function in order to get the plugin started.
' In addition to that, the host form must pass itself as an object
' to give the plugin access to the host applications objects and
' controls. So, the plugin will be able to modify/use any public
' objects on the host form...nice eh ?
On Error GoTo errtrap
    Load frmMain
                                   
    Set frmMain.objHost = FormX
    frmMain.Show vbModal            'VBModal is optional.  Remove
                                    'if you wish for the form to
                                    'be a non modal form
    Run = vbNullString
Exit Function
errtrap:
    Run = Err.Description
End Function

Public Function Identify() As String
'This function is one of the interface calls to the plugin. It allows the host to get
' a description of the plugin. This value could then be used
' in the host's menu.

  Identify = "Encryption Plugin version " & App.Major & "." & App.Minor & "." & App.Revision

End Function
