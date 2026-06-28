"""携帯番号表示VBAを「列全体を再表示」から「選択行だけパスワードで確認」に変更する。
一度だけローカルで実行する（Graph APIはVBAプロジェクトを操作できないため）。

実行手順:
  1. OneDrive上の ドライバー情報_営業用.xlsm をローカルにコピー（OneDrive同期フォルダ内の
     ファイルを直接 win32com で開くとアクセス拒否になることがあるため）。
  2. PATH をそのコピー先に設定してこのスクリプトを実行。
  3. 更新できたコピーをOneDriveのファイルに上書きコピーして戻す。
"""
import win32com.client as win32

PATH = r"C:\Users\1229\Desktop\AI関連－仮保存フォルダ\_vba_work.xlsm"  # 手順1のローカルコピー先に合わせて変更

NEW_VBA_CODE = '''
Sub ShowPhoneForActiveRow()
    Dim pwd As String
    Dim r As Long
    Dim phoneReal As String
    Dim caseNo As String

    r = ActiveCell.Row
    If r < 7 Then
        MsgBox "携帯番号を確認したい行のセルを選択してから実行してください。", vbExclamation
        Exit Sub
    End If

    phoneReal = Trim(ActiveSheet.Cells(r, "M").Value)
    If phoneReal = "" Then
        MsgBox "この行には携帯番号が登録されていません。", vbExclamation
        Exit Sub
    End If

    pwd = InputBox("パスワードを入力してください", "携帯番号の表示")
    If pwd = "" Then Exit Sub

    If pwd = "1229" Then
        caseNo = ActiveSheet.Cells(r, "A").Value
        MsgBox "案件No " & caseNo & " の携帯番号:" & vbCrLf & phoneReal, _
            vbInformation, "携帯番号（この行のみ表示）"
    Else
        MsgBox "パスワードが違います。", vbExclamation
    End If
End Sub
'''


def main() -> None:
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(PATH)
    try:
        vb_project = wb.VBProject
        target_module = None
        for component in vb_project.VBComponents:
            code_module = component.CodeModule
            if code_module.CountOfLines > 0:
                text = code_module.Lines(1, code_module.CountOfLines)
                if "ShowMobileNumbers" in text or "ShowPhoneForActiveRow" in text:
                    target_module = component
                    break
        if target_module is None:
            target_module = vb_project.VBComponents.Add(1)  # vbext_ct_StdModule

        code_module = target_module.CodeModule
        if code_module.CountOfLines > 0:
            code_module.DeleteLines(1, code_module.CountOfLines)
        code_module.AddFromString(NEW_VBA_CODE)

        ws = wb.Sheets("ドライバー情報")
        button_updated = False
        for shape in ws.Shapes:
            if shape.OnAction in ("ShowMobileNumbers", "ShowPhoneForActiveRow"):
                shape.OnAction = "ShowPhoneForActiveRow"
                shape.TextFrame.Characters().Text = "携帯番号を表示\n（選択した行だけ）"
                button_updated = True
                break
        if not button_updated:
            print("警告: 既存ボタンが見つかりませんでした。手動確認してください。")

        wb.Save()
        print("VBA更新完了: ShowPhoneForActiveRow に置き換えました。")
    finally:
        wb.Close(SaveChanges=False)
        excel.Quit()


if __name__ == "__main__":
    main()
