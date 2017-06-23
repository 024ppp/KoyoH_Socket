Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports SZKDLL_Orc
Imports SZKDLL_Orc.Class
Imports SZKDLL_Orc.DBACCESS

'TODO クラス名変更 DBに対する要求まとめクラスらしく
Public Class DBAction
    'クラス名
    Private Const C_CLASSNAME As String = "DBAction.vb"
    Private m_DT As DataTable = Nothing
    Private m_ErrMsg As String = ""
    Dim m_clsCom As New clsCom

    'データテーブルの取得
    Public Property rtnDT() As DataTable
        Get
            Return m_DT
        End Get
        Set(ByVal Value As DataTable)
        End Set
    End Property

    'エラーメッセージ
    Public Property ErrMsg() As String
        Get
            Return m_ErrMsg
        End Get
        Set(ByVal Value As String)
        End Set
    End Property

    '作業者取得
    Public Function getSagyoName(ByRef names As String) As Boolean
        Dim strSQL As String = ""
        Dim ds As New clsDsCtrl

        Try
            strSQL = " SELECT * "
            strSQL &= " FROM MT01 "
            strSQL &= " WHERE DELKBN = 0 "
            strSQL &= " ORDER BY TNTCOD "

            Call ds.Connect()
            Call ds.ExecuteSQL(strSQL, "")

            If ds.DataTable.Rows.Count <= 0 Then
                m_ErrMsg = "作業者名取得エラー。"
                Return False
            End If

            For Each dtRow As DataRow In ds.DataTable.Rows
                If Not names.Equals("") Then
                    names &= ","
                End If
                names &= dtRow.Item("TNTNAM").ToString
            Next

            Return True
        Catch ex As Exception
            m_ErrMsg = ex.Message
            Return False
        Finally
            Call ds.Dispose()
        End Try
    End Function

    'Vコンが存在するかチェックして結果を返す
    Public Function checkVkonName(ByVal sKiknam As String) As Boolean
        Dim strSQL As String = ""
        Dim ds As New clsDsCtrl

        Try
            strSQL = " SELECT * "
            strSQL &= " FROM MV01 "
            strSQL &= " WHERE 1 = 1 "
            If sKiknam <> "" Then
                strSQL &= " AND KIKNAM = '" & sKiknam & "' "
            End If

            Call ds.Connect()
            Call ds.ExecuteSQL(strSQL, "")

            If ds.DataTable.Rows.Count <= 0 Then
                m_ErrMsg = "Vコンが存在しません。"
                Return False
            End If

            Return True
        Catch ex As Exception
            m_ErrMsg = ex.Message
            Return False
        Finally
            Call ds.Dispose()
        End Try
    End Function

    'MV01更新
    Public Function Update(ByVal strUpdText As String) As Boolean
        Dim strSQL As String = ""
        Dim arrUpdText As String() = Nothing
        Dim sMesy As String
        Dim sKiknam As String

        Dim ds As New clsDsCtrl

        Try
            arrUpdText = strUpdText.Split(",")
            sKiknam = arrUpdText(0)
            sMesy = arrUpdText(1)

            'SQL文作成
            strSQL &= " UPDATE MV01 SET "
            strSQL &= " MESY = '" & sMesy & "'"
            strSQL &= " ,UPDYMD = '" & Now.ToString("yyyyMMdd") & "' "
            strSQL &= " ,UPDHMS = '" & Now.ToString("HHmmss") & "' "
            strSQL &= " WHERE "
            strSQL &= " KIKNAM = '" & sKiknam & "' "

            Call ds.Connect()
            Call ds.ExecuteSQL(strSQL, "")

            Return True
        Catch ex As Exception
            m_ErrMsg = ex.Message
            Return False
        End Try
    End Function

    

End Class
