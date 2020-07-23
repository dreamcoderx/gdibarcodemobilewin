Imports System.IO
Imports System.Xml
Imports System.Collections.Specialized

Public Class Settings
    Private Shared m_settings As NameValueCollection
    Private Shared m_settingsPath As String


    Shared Sub New()
        'Get the path of the settings file.
        m_settingsPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.GetName().CodeBase)
        m_settingsPath &= "\Settings.xml"

        If Not File.Exists(m_settingsPath) Then
            Throw New FileNotFoundException(m_settingsPath & " could not be found.")
        End If

        Dim xdoc As New XmlDocument
        xdoc.Load(m_settingsPath)
        Dim root As XmlElement = xdoc.DocumentElement
        Dim nodeList As XmlNodeList = root.ChildNodes.Item(0).ChildNodes

        'Add settings to the NameValueCollection
        m_settings = New NameValueCollection
        For Each item As XmlNode In nodeList
            If item.Name = "add" Then
                m_settings.Add(item.Attributes("key").Value, item.Attributes("value").Value)
            End If
        Next
    End Sub

    Public Shared Function [Get](ByVal key As String) As String
        Return m_settings.Get(key)
    End Function

    Public Shared Sub [Set](ByVal key As String, ByVal value As String)
        m_settings.Set(key, value)
    End Sub

    Public Shared Sub Update()
        Dim tw As New XmlTextWriter(m_settingsPath, System.Text.UTF8Encoding.UTF8)

        tw.WriteStartDocument()
        tw.WriteStartElement("configuration")
        tw.WriteStartElement("appSettings")

        For Each key As String In m_settings.AllKeys
            tw.WriteStartElement("add")
            tw.WriteStartAttribute("key", String.Empty)
            tw.WriteRaw(key)
            tw.WriteEndAttribute()

            tw.WriteStartAttribute("value", String.Empty)
            tw.WriteRaw(m_settings.Item(key))
            tw.WriteEndAttribute()
            tw.WriteEndElement()
        Next

        tw.WriteEndElement()
        tw.WriteEndElement()

        tw.Close()
    End Sub
End Class
