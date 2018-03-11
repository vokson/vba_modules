Attribute VB_Name = "M_VBA_Library_List_v2"
Option Explicit

Public Function getListOfRequiredModules() As Dictionary
    Dim dic As New Dictionary

    ' ФОРМАТ ВЕРСИИ
    ' "AB.B.B"
    ' A = ["", "=", ">", "<", ">=", "<="]
    ' B = "*" либо комбинация цифр ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9"]
    ' Зависимые модули загружаются автоматически согласно правилам, указанным в package.json
    ' каждого модуля. При появлении новой зависимости, выбирается максимальная версия
    ' модуля, подходящего под правило. Эта версия проверяется также, если зависимость
    ' появляется повторно. Если версия не удовлетворяет новой зависимости, выдается ошибка
    
    dic.Item("C_Soil_Database") = "1.1.0"
    
    Set getListOfRequiredModules = dic
End Function
