Attribute VB_Name = "periodificacion"
'
' Periodificacion del suministro electrico
'
' La Circular 3/2020 del CNMC tiene como objeto el establecimiento de la metodología para el cálculo anual de los precios de los peajes de acceso a las redes de transporte y distribución de electricidad.
' A partir del 4 de abril del 2021 y la entrada en vigor de la metodología de cálculo, cambian los periodos de potencia y energia de los suministros.
' Las funciones definidas en la presente macro devuelven el codigo de periodopara un argumento de fecha y hora.
'
' periodo_30TD6XTD(_arg_): devuelve un entero de 1 a 6 que corresponde al número del periodo de la hora introducida como argumento (_arg_) para un suministro conectado con nivel de tensión 3.0TD a 6.xTD
' periodo_energia_20TD(_arg_): devuelve un entero de 1 a 3 (1:punta, 2:llano, 3:valle) que corresponde al número del periodo de energia de la hora introducida como argumento (_arg_) para un suministro conectado con nivel de tensión 2.0TD
' periodo_potencia_20TD(_arg_): devuelve un entero 1 o 3 (1:punta, 3:valle) que corresponde al número del periodo de potencia de la hora introducida como argumento (_arg_) para un suministro conectado con nivel de tensión 2.0TD
'
' Las funciones estan configuradas para el sistema peninsular y se deben ajustar para los otros sistemas.
' Las funciones se entregan "AS IS" sin garantias ni responsabilidades.


' -------------- LISTA DE DIAS FESTIVOS -----------------
Dim festivos As Variant
Sub dias_festivos()
    'Incluir la lista de los dias festivos en formato MM-dd.
    festivos = Array("01-01", "01-06", "05-01", "08-15", "10-12", "11-01", "12-06", "12-08", "12-25")
    'Se consideran a estos efectos como días festivos el 6 de enero y los de ámbito nacional, definidos como tales en el calendario oficial del año correspondiente, con exclusión tanto de los festivos sustituibles como de los que no tienen fecha fija.
    'vease BOE-A-2020-13343 https://www.boe.es/eli/es/res/2020/10/28/(1)
End Sub


' -------------- PERIODOS PARA NIVEL DE TENSION DE 3.0TD a 6.XTD -----------------
Function periodo_30TD6XTD(timeStamp As Date)
    Dim periodo As Integer
    
    Dim hora As Integer
    hora = hour(timeStamp)
    
    Dim dia As Integer
    dia = Weekday(timeStamp)
    
    Dim mes As Integer
    mes = Month(timeStamp)
    
    'El periodo por defecto es 6
    periodo = 6
        
    'Modificacion del periodo segun la temporada
    Select Case mes
        'Si es un mes de temporada Baja no se modifica el periodo
        Case 4, 5, 10
            periodo = periodo
        'Si es un mes de temporada Media el periodo baja de 1
        Case 6, 8, 9
            periodo = periodo - 1
        'Si es un mes de temporada Media-Alta el periodo baja de 2
        Case 3, 11
            periodo = periodo - 2
        'Si es un mes de temporada Alta el periodo baja de 3
        Case 1, 2, 7, 12
            periodo = periodo - 3
    End Select
        
    'Modificacion del periodo segun el periodo horario
    Select Case hora
        'Si es una hora dentro del periodo de 0h a 8h el periodo es 6
        Case 0, 1, 2, 3, 4, 5, 6, 7
            periodo = 6
        'Si es una hora dentro de los periodos de 8h a 9h o de 14h a 18h o de 22h a 0h el periodo baja de 1
        Case 8, 14, 15, 16, 17, 22, 23
            periodo = periodo - 1
        'Si es una hora dentro de los periodos de 9h a 14h o de 18h a 22h el periodo baja de 2
        Case 9, 10, 11, 12, 13, 18, 19, 20, 21
            periodo = periodo - 2
    End Select

    'Si es un sabado o un domingo el periodo es 6
    Select Case dia
        Case 1, 7
        periodo = 6
    End Select
    
    'Si es un dia festivo el periodo es 6
    dias_festivos
    For Each x In festivos
        If x = Format(timeStamp, "MM-dd") Then
            periodo = 6
        End If
    Next
    
    periodo_30TD6XTD = periodo
End Function


' -------------- PERIODOS DE ENERGIA PARA NIVEL DE TENSION DE 2.0TD -----------------
Function periodo_energia_20TD(timeStamp As Date)
    Dim periodo As Integer
    
    Dim hora As Integer
    hora = hour(timeStamp)
    
    Dim dia As Integer
    dia = Weekday(timeStamp)
    
    Dim mes As Integer
    mes = Month(timeStamp)
    
    'El periodo por defecto es 3
    periodo = 3
        
    'Modificacion del periodo segun el periodo horario
    Select Case hora
        'Si es una hora dentro del periodo de 0h a 8h el periodo es 3
        Case 0, 1, 2, 3, 4, 5, 6, 7
            periodo = 3
        'Si es una hora dentro de los periodos de 8h a 10h o de 14h a 18h o de 22h a 0h el periodo es 2
        Case 8, 9, 14, 15, 16, 17, 22, 23
            periodo = 2
        'Si es una hora dentro de los periodos de 10h a 14h o de 18h a 22h el periodo es 1
        Case 10, 11, 12, 13, 18, 19, 20, 21
            periodo = 1
    End Select

    'Si es un sabado o un domingo el periodo es 3
    Select Case dia
        Case 1, 7
        periodo = 3
    End Select
    
    'Si es un dia festivo el periodo es 3
    dias_festivos
    For Each x In festivos
        If x = Format(timeStamp, "yyyy-MM-dd") Then
            periodo = 3
        End If
    Next
    
    periodo_energia_20TD = periodo
End Function


' -------------- PERIODOS DE POTENCIA PARA NIVEL DE TENSION DE 2.0TD -----------------
Function periodo_potencia_20TD(timeStamp As Date)
    Dim periodo As Integer
    
    Dim hora As Integer
    hora = hour(timeStamp)
    
    Dim dia As Integer
    dia = Weekday(timeStamp)
    
    Dim mes As Integer
    mes = Month(timeStamp)
    
    'El periodo por defecto es 3
    periodo = 3
        
    'Modificacion del periodo segun el periodo horario
    Select Case hora
        'Si es una hora dentro del periodo de 0h a 8h el periodo es 3
        Case 0, 1, 2, 3, 4, 5, 6, 7
            periodo = 3
        'Si es una hora fuera del periodo de 0h a 8h el periodo es 1
        Case Else
            periodo = 1
    End Select

    'Si es un sabado o un domingo el periodo es 3
    Select Case dia
        Case 1, 7
        periodo = 3
    End Select
    
    'Si es un dia festivo el periodo es 3
    dias_festivos
    For Each x In festivos
        If x = Format(timeStamp, "yyyy-MM-dd") Then
            periodo = 3
        End If
    Next
    
    periodo_potencia_20TD = periodo
End Function
