# periodificación
Excel macro building specific function to calculate periods as specified in Circular 3/2020 of spanish CNMC

La Circular 3/2020 del CNMC tiene como objeto el establecimiento de la metodología para el cálculo anual de los precios de los peajes de acceso a las redes de transporte y distribución de electricidad.<br>
A partir del **4 de abril del 2021** y la entrada en vigor de la nueva metodología de cálculo, **cambian los periodos de potencia y de energia** de los suministros.<br>
Las funciones definidas en la presente macro devuelven el codigo de periodo para un argumento de fecha y hora.


>## Instalación
> Para poder usar las funciones solo es necesario descargar el archivo _periodificacion.bas_ e insatalar la macro.
> 
><img src="https://user-images.githubusercontent.com/73427462/109394693-91668580-7928-11eb-9ed4-0d2fd13abaa2.gif" data-canonical-src="https://user-images.githubusercontent.com/73427462/109394693-91668580-7928-11eb-9ed4-0d2fd13abaa2.gif" width="250" height="200" />
>
> Guardar el archivo Excel en formato .xlsm permite conservar la marco en el archivo y por lo tanto las funciones.


>## Utilización
>Una vez instalada la macro se pueden usar las funciones como cualquier otra función de Excel
>
>![Utilizacion](https://user-images.githubusercontent.com/73427462/109396142-713ac480-7930-11eb-80f8-67f7fe98dba4.gif)
>

>## Funciones disponibles
>Una vez instalada la macro se pueden usar la funciones como cualquier otra función de Excel
>
>Con _arg_ como una fecha con hora.
>
>> #### periodo_30TD6XTD(_arg_)
>> devuelve un entero de 1 a 6 que corresponde al número del periodo de la hora introducida como argumento (_arg_) para un suministro conectado con nivel de tensión 3.0TD a 6.xTD
>
>> #### periodo_energia_20TD(_arg_)
>> devuelve un entero de 1 a 3 (1:punta, 2:llano, 3:valle) que corresponde al número del periodo de energia de la hora introducida como argumento (_arg_) para un suministro conectado con nivel de tensión 2.0TD
>
>> #### periodo_potencia_20TD(_arg_)
>> devuelve un entero 1 o 3 (1:punta, 3:valle) que corresponde al número del periodo de potencia de la hora introducida como argumento (_arg_) para un suministro conectado con nivel de tensión 2.0TD
>

>## Personalización
> Las funciones están configuradas para calcular los periodos del ***sistema Peninsular***.<br>
> Se puede personalizar las funciones para que devuelvan el periodo correspondiendo a otro sistema ***(Canarias, Illes Balears, Ceuta o Melilla)***, modificando los parámetros de meses según lo especificado en el BOE.
>``` 
>    'Modificacion del periodo segun la temporada
>    Select Case mes
>        'Si es un mes de temporada Baja no se modifica el periodo
>        Case 4, 5, 10
>            periodo = periodo
>        'Si es un mes de temporada Media el periodo baja de 1
>        Case 6, 8, 9
>            periodo = periodo - 1
>        'Si es un mes de temporada Media-Alta el periodo baja de 2
>        Case 3, 11
>            periodo = periodo - 2
>        'Si es un mes de temporada Alta el periodo baja de 3
>        Case 1, 2, 7, 12
>            periodo = periodo - 3
>    End Select
>```
>


>## Fuente
> https://www.boe.es/buscar/doc.php?id=BOE-A-2020-1066
