Attribute VB_Name = "mdlListConverter"
'@Folder("main")
Option Explicit

Sub ConvertListNumbers()
    '�ӏ������̒i���ԍ������e�L�X�g�ɕϊ�
    ActiveDocument.ConvertNumbersToText wdNumberParagraph
    With Application.Options
        '���̓I�[�g�t�H�[�}�b�g/���͒��Ɏ����ŕύX���鍀��/�s�̎n�܂�̃X�y�[�X���������ɕύX����
        .AutoFormatApplyFirstIndents = True
        '���̓I�[�g�t�H�[�}�b�g/���͒��Ɏ����ŏ����ݒ肷�鍀��/�ӏ������i�s�������j
        .AutoFormatAsYouTypeApplyBulletedLists = False
        '���̓I�[�g�t�H�[�}�b�g/���͒��Ɏ����ŏ����ݒ肷�鍀��/�ӏ������i�i���ԍ��j
        .AutoFormatAsYouTypeApplyNumberedLists = False
        '�I�[�g�t�H�[�}�b�g/�����œK�؂ȃX�^�C����ݒ肷��ӏ�/�ӏ������i�s�������j
        .AutoFormatApplyLists = False
        '�I�[�g�t�H�[�}�b�g/�����œK�؂ȃX�^�C����ݒ肷��ӏ�/���X�g�̃X�^�C��
        .AutoFormatApplyBulletedLists = False
    End With
    
End Sub
