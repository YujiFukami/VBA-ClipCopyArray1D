Attribute VB_Name = "ModClipCopyArray1D"
Option Explicit

'ClipCopyArray1D   �E�E�E���ꏊ�FFukamiAddins3.ModArray    
'ClipboardCopy     �E�E�E���ꏊ�FFukamiAddins3.ModClipboard
'CheckArray1D      �E�E�E���ꏊ�FFukamiAddins3.ModArray    
'CheckArray1DStart1�E�E�E���ꏊ�FFukamiAddins3.ModArray    

'------------------------------


'�z��̏����֌W�̃v���V�[�W��

'------------------------------



'------------------------------


Public Sub ClipCopyArray1D(Array1D)
'1�����z���ϐ��錾�p�̃e�L�X�g�f�[�^�ɕϊ����āA�N���b�v�{�[�h�ɃR�s�[����
'20210805
    
    '�����`�F�b�N
    Call CheckArray1D(Array1D, "Array1D")
    Call CheckArray1DStart1(Array1D, "Array1D")
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    N = UBound(Array1D, 1)
    
    Dim TmpValue
    Dim Output$
    
    Output = String(3, Chr(9)) & "Array("
    For I = 1 To N
        
        TmpValue = Array1D(I)
        If IsNumeric(TmpValue) Then
            Output = Output & TmpValue
        Else
            Output = Output & """" & TmpValue & """"
        End If
        
        If I < N Then
            Output = Output & ","
        Else
            Output = Output & ")"
        End If
        
    Next I
    
    Output = "Application.Transpose(Application.Transpose( _" & vbLf & Output & " _" & vbLf & String(3, Chr(9)) & "))"
    
    Call ClipboardCopy(Output, True)
    
End Sub

Private Sub ClipboardCopy(ByVal InputClipText, Optional MessageIrunaraTrue As Boolean = False)
'���̓e�L�X�g���N���b�v�{�[�h�Ɋi�[
'�z��Ȃ�Η������Tab�킯�A�s���������s����B
'20210719�쐬
    
    '���͂����������z�񂩁A�z��̏ꍇ��1�����z�񂩁A2�����z�񂩔���
    Dim HairetuHantei%
    Dim Jigen1%, Jigen2%
    If IsArray(InputClipText) = False Then
        '���͈������z��łȂ�
        HairetuHantei = 0
    Else
        On Error Resume Next
        Jigen2 = UBound(InputClipText, 2)
        On Error GoTo 0
        
        If Jigen2 = 0 Then
            HairetuHantei = 1
        Else
            HairetuHantei = 2
        End If
    End If
    
    '�N���b�v�{�[�h�Ɋi�[�p�̃e�L�X�g�ϐ����쐬
    Dim Output$
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    
    If HairetuHantei = 0 Then '�z��łȂ��ꍇ
        Output = InputClipText
    ElseIf HairetuHantei = 1 Then '1�����z��̏ꍇ
    
        If LBound(InputClipText, 1) <> 1 Then '�ŏ��̗v�f�ԍ���1�o�Ȃ��ꍇ�͍ŏ��̗v�f�ԍ���1�ɂ���
            InputClipText = Application.Transpose(Application.Transpose(InputClipText))
        End If
        
        N = UBound(InputClipText, 1)
        
        Output = ""
        For I = 1 To N
            If I = 1 Then
                Output = InputClipText(I)
            Else
                Output = Output & vbLf & InputClipText(I)
            End If
            
        Next I
    ElseIf HairetuHantei = 2 Then '2�����z��̏ꍇ
        
        If LBound(InputClipText, 1) <> 1 Or LBound(InputClipText, 2) <> 1 Then
            InputClipText = Application.Transpose(Application.Transpose(InputClipText))
        End If
        
        N = UBound(InputClipText, 1)
        M = UBound(InputClipText, 2)
        
        Output = ""
        
        For I = 1 To N
            For J = 1 To M
                If J < M Then
                    Output = Output & InputClipText(I, J) & Chr(9)
                Else
                    Output = Output & InputClipText(I, J)
                End If
                
            Next J
            
            If I < N Then
                Output = Output & Chr(10)
            End If
        Next I
    End If
    
    
    '�N���b�v�{�[�h�Ɋi�['�Q�l https://www.ka-net.org/blog/?p=7537
    With CreateObject("Forms.TextBox.1")
        .MultiLine = True
        .Text = Output
        .SelStart = 0
        .SelLength = .TextLength
        .Copy
    End With

    '�i�[�����e�L�X�g�ϐ������b�Z�[�W�\��
    If MessageIrunaraTrue Then
        MsgBox ("�u" & Output & "�v" & vbLf & _
                "���N���b�v�{�[�h�ɃR�s�[���܂����B")
    End If
    
End Sub

Private Sub CheckArray1D(InputArray, Optional HairetuName$ = "�z��")
'���͔z��1�����z�񂩂ǂ����`�F�b�N����
'20210804

    Dim Dummy%
    On Error Resume Next
    Dummy = UBound(InputArray, 2)
    On Error GoTo 0
    If Dummy <> 0 Then
        MsgBox (HairetuName & "��1�����z�����͂��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Private Sub CheckArray1DStart1(InputArray, Optional HairetuName$ = "�z��")
'����1�����z��̊J�n�ԍ���1���ǂ����`�F�b�N����
'20210804

    If LBound(InputArray, 1) <> 1 Then
        MsgBox (HairetuName & "�̊J�n�v�f�ԍ���1�ɂ��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub


