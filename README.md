# XLSX2CSV

Создает csv из xlsx файла.


## Тесты

На входе: xlsx файл с числами разных типов, форматов

На выходе: полученный csv файл

* строки 2-6 - int с различными форматами ячеек, которые предлагает libre office calc (general, number, date, time, scientific)
C scientific, date, time все в порядке, а number и general, как ни странно, выдают что-то странное
* строки 8-10 - double разных порядков
* строка 11 - int с большим числом знаков - все в порядке
* строка 13 - точка вместо запятой - все в порядке
* строка 14 - добавление пробелов между цифрами (на скрине не видно) - все в порядке

![from](/img/xl.PNG)  ====>  ![to](img/csv.PNG)
