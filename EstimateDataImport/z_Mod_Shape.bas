Attribute VB_Name = "z_Mod_Shape"
Function ShapeInRange(rng As Range, _
                Optional iRed As Long = 255, _
                Optional iGreen As Long = 0, _
                Optional iBlue As Long = 0, _
                Optional FillVisible As MsoTriState = msoTrue, _
                Optional LineVisible As MsoTriState = msoTrue, _
                Optional Transparent As Double = 0.95, _
                Optional LineWeight As Double = 0.5, _
                Optional DashType As MsoLineDashStyle = msoLineDash, _
                Optional ShapeType As MsoAutoShapeType = msoShapeRectangle, _
                Optional ActivateSheet As Boolean = True) As Shape

'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'���� �� ���� �� ��ó�� �ݵ�� ����ؾ� �մϴ�.
'
'�� ShapeInRange ��ɹ�
'�� ���õ� ���� �ȿ� ������ �����մϴ�.
'�� �����
''�Ʒ� ��ɹ��� ������ ���ȿ� �ٿ��ֱ� �� ��, '����Ʈ�ڽ�' �� ���� ������ ����Ʈ�ڽ� �̸����� �����մϴ�.
''----------------------------------------------------------------------------------------------------
'Dim Shp As Shape
'Set Shp = ShapeInRange(Range("A1"))
''-----------------------------------------------------------------------------------------------------
'�� �μ� ����
'_____________Rng               : ������ ������ �����Դϴ�.
'_____________iRed               : [�����μ�] ������ ������ RGB, R ���Դϴ�. �⺻���� 255 �Դϴ�.
'_____________iGreen           : [�����μ�] ������ ������ RGB, G ���Դϴ�. �⺻���� 0 �Դϴ�.
'_____________iBlue              : [�����μ�] ������ ������ RGB, B ���Դϴ�. �⺻���� 0 �Դϴ�.
'_____________FillVisible       : [�����μ�] ä��� �����Դϴ�. �⺻���� TRUE �Դϴ�.
'_____________LineVisible     : [�����μ�] ������ �����Դϴ�. �⺻���� TRUE �Դϴ�.
'_____________Transparent  : [�����μ�] ä��� �����Դϴ�. �⺻���� 0.25 �Դϴ�.
'_____________LineWeight   : [�����μ�] ������ �β��Դϴ�. �⺻���� 0.5 �Դϴ�.
'_____________DashType      : [�����μ�] ������ ��Ÿ���Դϴ�. �⺻���� �����Դϴ�.
'_____________ShapeType    : [�����μ�] ���� ����Դϴ�. �⺻���� ���簢���Դϴ�.
'_____________AvtiveSheet   : [�����μ�] �������� �� ���Ե� ��Ʈ Ȱ��ȭ�����Դϴ�. �⺻���� True �Դϴ�.
'###############################################################

Dim Shp As Shape
Dim WS As Worksheet

Set WS = rng.Parent

With rng
    Set Shp = WS.Shapes.AddShape(ShapeType, .Left, .top, .Width, .Height)
End With

With Shp
    With .Fill
        .Visible = FillVisible
        .ForeColor.RGB = RGB(iRed, iGreen, iBlue)
        .Transparency = Transparent
    End With
    With .Line
        .Visible = LineVisible
        .ForeColor.RGB = RGB(iRed, iGreen, iBlue)
        .Weight = LineWeight
        .DashStyle = DashType
    End With
End With

Set ShapeInRange = Shp
If ActivateSheet = True Then Shp.Parent.Activate

End Function

