---
title: Вызов встроенных функций листов Excel с помощью API JavaScript для Excel
description: 'Узнайте, как вызывать встроенные функции Excel `VLOOKUP` `SUM` таблицы, такие как Excel API JavaScript.'
ms.date: 12/19/2019
ms.localizationpriority: medium
---

# <a name="call-built-in-excel-worksheet-functions"></a>Вызов встроенных функций листов Excel

В этой статье рассказывается, как вызывать встроенные функции листов Excel, такие как `VLOOKUP` и `SUM`, с помощью API JavaScript для Excel. В ней также представлен полный список встроенных функций листов Excel, которые можно вызывать с помощью API JavaScript для Excel.

> [!NOTE]
> Сведения о том, как создавать *пользовательские функции* в Excel с помощью API JavaScript для Excel, см. в статье [Создание пользовательских функций в Excel](custom-functions-overview.md).

## <a name="calling-a-worksheet-function"></a>Вызов функции листа

В приведенном ниже фрагменте кода показано, как вызвать функцию листа, где `sampleFunction()`— это заполнитель, который следует заменить на имя вызываемой функции и необходимые ей входные параметры. Свойство `value` объекта `FunctionResult` , возвращаемого функцией таблицы, содержит результат указанной функции. Как показано в этом примере `load` `value` `FunctionResult` , перед чтением необходимо свойство объекта. В этом примере результат выполнения функции просто записывается в консоль.

```js
var functionResult = context.workbook.functions.sampleFunction();
functionResult.load('value');
return context.sync()
    .then(function () {
        console.log('Result of the function: ' + functionResult.value);
    });
```

> [!TIP]
> В разделе [Поддерживаемые функции листов](#supported-worksheet-functions) в этой статье представлен список функций, которые можно вызывать с помощью API JavaScript для Excel.

## <a name="sample-data"></a>Образец данных

На приведенном ниже изображении показана таблица на листе Excel, содержащая данные о продажах различных инструментов в течение трех месяцев. Каждое число в таблице представляет количество единиц того или иного инструмента, проданных за определенный месяц. В последующих примерах показано, как применить к этим данным встроенные функции листов.

![Снимок экрана данных о продажах Excel для Hammer, Wrench и Saw в ноябре, декабре и январе.](../images/worksheet-functions-chaining-results.jpg)

## <a name="example-1-single-function"></a>Пример 1. Одна функция

В приведенном ниже примере кода к вышеописанному примеру данных применяется функция `VLOOKUP`, чтобы определить количество гаечных ключей, проданных в ноябре.

```js
Excel.run(function (context) {
    var range = context.workbook.worksheets.getItem("Sheet1").getRange("A1:D4");
    var unitSoldInNov = context.workbook.functions.vlookup("Wrench", range, 2, false);
    unitSoldInNov.load('value');

    return context.sync()
        .then(function () {
            console.log(' Number of wrenches sold in November = ' + unitSoldInNov.value);
        });
}).catch(errorHandlerFunction);
```

## <a name="example-2-nested-functions"></a>Пример 2. Вложенные функции

В приведенном ниже примере кода к вышеописанному примеру данных применяется функция `VLOOKUP`, чтобы определить количество гаечных ключей, проданных в ноябре и декабре, а затем применяется функция `SUM`, чтобы вычислить общее число гаечных ключей, проданных за эти два месяца.

Как показано в этом примере, если один или несколько вызовов функций вложены в вызов другой функции, то выполнять операцию `load` нужно только с окончательным результатом, который впоследствии потребуется прочитать (в этом примере — `sumOfTwoLookups`). Все промежуточные результаты (в этом примере — результат выполнения каждой функции `VLOOKUP`) будут вычислены и использованы для вычисления окончательного результата.

```js
Excel.run(function (context) {
    var range = context.workbook.worksheets.getItem("Sheet1").getRange("A1:D4");
    var sumOfTwoLookups = context.workbook.functions.sum(
        context.workbook.functions.vlookup("Wrench", range, 2, false),
        context.workbook.functions.vlookup("Wrench", range, 3, false)
    );
    sumOfTwoLookups.load('value');

    return context.sync()
        .then(function () {
            console.log(' Number of wrenches sold in November and December = ' + sumOfTwoLookups.value);
        });
}).catch(errorHandlerFunction);
```

## <a name="supported-worksheet-functions"></a>Поддерживаемые функции листов

Ниже перечислены встроенные функции листов Excel, которые можно вызывать с помощью API JavaScript для Excel

| Функция | Описание |
|:---------------|:-----------|
| <a href="https://support.microsoft.com/office/3420200f-5628-4e8c-99da-c99d7c87713c" target="_blank">Функция ABS</a> | Возвращает абсолютное значение числа |
| <a href="https://support.microsoft.com/office/fe45d089-6722-4fb3-9379-e1f911d8dc74" target="_blank">Функция НАКОПДОХОД</a> | Возвращает накопленный процент по ценной бумаге с периодической выплатой процентов |
| <a href="https://support.microsoft.com/office/f62f01f9-5754-4cc4-805b-0e70199328a7" target="_blank">Функция НАКОПДОХОДПОГАШ</a> | Возвращает накопленный процент по ценной бумаге, процент по которой выплачивается в срок погашения |
| <a href="https://support.microsoft.com/office/cb73173f-d089-4582-afa1-76e5524b5d5b" target="_blank">Функция ACOS</a> | Возвращает арккосинус числа |
| <a href="https://support.microsoft.com/office/e3992cc1-103f-4e72-9f04-624b9ef5ebfe" target="_blank">Функция ACOSH</a> | Возвращает обратный гиперболический косинус числа (ареакосинус) |
| <a href="https://support.microsoft.com/office/dc7e5008-fe6b-402e-bdd6-2eea8383d905" target="_blank">Функция ACOT</a> | Возвращает арккотангенс числа |
| <a href="https://support.microsoft.com/office/cc49480f-f684-4171-9fc5-73e4e852300f" target="_blank">Функция ACOTH</a> | Возвращает гиперболический арккотангенс числа |
| <a href="https://support.microsoft.com/office/a14d0ca1-64a4-42eb-9b3d-b0dededf9e51" target="_blank">Функция АМОРУМ</a> | Возвращает величину амортизации для каждого учетного периода, используя коэффициент амортизации |
| <a href="https://support.microsoft.com/office/7d417b45-f7f5-4dba-a0a5-3451a81079a8" target="_blank">Функция АМОРУВ</a> | Возвращает величину амортизации для каждого учетного периода |
| <a href="https://support.microsoft.com/office/5f19b2e8-e1df-4408-897a-ce285a19e9d9" target="_blank">Функция И</a> | Возвращает значение `TRUE`, если все аргументы имеют значение ИСТИНА |
| <a href="https://support.microsoft.com/office/9a8da418-c17b-4ef9-a657-9370a30a674f" target="_blank">Функция АРАБСКОЕ</a> | Преобразует римское число в арабское |
| <a href="https://support.microsoft.com/office/8392ba32-7a41-43b3-96b0-3695d2ec6152" target="_blank">Функция ОБЛАСТИ</a> | Возвращает количество областей в ссылке |
| <a href="https://support.microsoft.com/office/0b6abf1c-c663-4004-a964-ebc00b723266" target="_blank">Функция ASC</a> | Преобразует полноширинные (двухбайтовые) английские буквы или знаки катакана в строке символов в полуширинные (однобайтовые) символы |
| <a href="https://support.microsoft.com/office/81fb95e5-6d6f-48c4-bc45-58f955c6d347" target="_blank">Функция ASIN</a> | Возвращает арксинус числа |
| <a href="https://support.microsoft.com/office/4e00475a-067a-43cf-926a-765b0249717c" target="_blank">Функция ASINH</a> | Возвращает обратный гиперболический синус числа (ареасинус) |
| <a href="https://support.microsoft.com/office/50746fa8-630a-406b-81d0-4a2aed395543" target="_blank">Функция ATAN</a> | Возвращает арктангенс числа |
| <a href="https://support.microsoft.com/office/c04592ab-b9e3-4908-b428-c96b3a565033" target="_blank">Функция ATAN2</a> | Возвращает арктангенс для заданных координат x и y |
| <a href="https://support.microsoft.com/office/3cd65768-0de7-4f1d-b312-d01c8c930d90" target="_blank">Функция ATANH</a> | Возвращает обратный гиперболический тангенс числа (ареатангенс) |
| <a href="https://support.microsoft.com/office/58fe8d65-2a84-4dc7-8052-f3f87b5c6639" target="_blank">Функция СРОТКЛ</a> | Возвращает среднее арифметическое абсолютных отклонений значений от их среднего |
| <a href="https://support.microsoft.com/office/047bac88-d466-426c-a32b-8f33eb960cf6" target="_blank">Функция СРЗНАЧ</a> | Возвращает среднее арифметическое аргументов |
| <a href="https://support.microsoft.com/office/f5f84098-d453-4f4c-bbba-3d2c66356091" target="_blank">Функция СРЗНАЧА</a> | Возвращает среднее значение аргументов (включая числовые, текстовые и логические) |
| <a href="https://support.microsoft.com/office/faec8e2e-0dec-4308-af69-f5576d8ac642" target="_blank">Функция СРЗНАЧЕСЛИ</a> | Возвращает среднее арифметическое всех ячеек в диапазоне, соответствующих определенному условию |
| <a href="https://support.microsoft.com/office/48910c45-1fc0-4389-a028-f7c5c3001690" target="_blank">Функция СРЗНАЧЕСЛИМН</a> | Возвращает среднее арифметическое всех ячеек, соответствующих нескольким условиям |
| <a href="https://support.microsoft.com/office/5ba4d0b4-abd3-4325-8d22-7a92d59aab9c" target="_blank">Функция БАТТЕКСТ</a> | Преобразует число в текст, используя денежный формат ß (бат) |
| <a href="https://support.microsoft.com/office/2ef61411-aee9-4f29-a811-1c42456c6342" target="_blank">Функция ОСНОВАНИЕ</a> | Преобразует число в текстовое представление с указанным основанием системы счисления |
| <a href="https://support.microsoft.com/office/8d33855c-9a8d-444b-98e0-852267b1c0df" target="_blank">Функция БЕССЕЛЬ.I</a> | Возвращает модифицированную функцию Бесселя In(x) |
| <a href="https://support.microsoft.com/office/839cb181-48de-408b-9d80-bd02982d94f7" target="_blank">Функция БЕССЕЛЬ.J</a> | Возвращает функцию Бесселя Jn(x) |
| <a href="https://support.microsoft.com/office/606d11bc-06d3-4d53-9ecb-2803e2b90b70" target="_blank">Функция БЕССЕЛЬ.K</a> | Возвращает модифицированную функцию Бесселя Kn(x) |
| <a href="https://support.microsoft.com/office/f3a356b3-da89-42c3-8974-2da54d6353a2" target="_blank">Функция БЕССЕЛЬ.Y</a> | Возвращает функцию Бесселя Yn(x) |
| <a href="https://support.microsoft.com/office/11188c9c-780a-42c7-ba43-9ecb5a878d31" target="_blank">Функция БЕТА.РАСП</a> | Возвращает функцию интегрального бета-распределения |
| <a href="https://support.microsoft.com/office/e84cb8aa-8df0-4cf6-9892-83a341d252eb" target="_blank">Функция БЕТА.ОБР</a> | Возвращает обратную функцию к интегральной функции указанного бета-распределения |
| <a href="https://support.microsoft.com/office/63905b57-b3a0-453d-99f4-647bb519cd6c" target="_blank">Функция ДВ.В.ДЕС</a> | Преобразует двоичное число в десятичное |
| <a href="https://support.microsoft.com/office/0375e507-f5e5-4077-9af8-28d84f9f41cc" target="_blank">Функция ДВ.В.ШЕСТН</a> | Преобразует двоичное число в шестнадцатеричное |
| <a href="https://support.microsoft.com/office/0a4e01ba-ac8d-4158-9b29-16c25c4c23fd" target="_blank">Функция ДВ.В.ВОСЬМ</a> | Преобразует двоичное число в восьмеричное |
| <a href="https://support.microsoft.com/office/c5ae37b6-f39c-4be2-94c2-509a1480770c" target="_blank">Функция БИНОМ.РАСП</a> | Возвращает вероятность биномиального распределения отдельного условия |
| <a href="https://support.microsoft.com/office/17331329-74c7-4053-bb4c-6653a7421595" target="_blank">Функция БИНОМ.РАСП.ДИАП</a> | Возвращает вероятность получения определенного результата испытания с помощью биномиального распределения |
| <a href="https://support.microsoft.com/office/80a0370c-ada6-49b4-83e7-05a91ba77ac9" target="_blank">Функция БИНОМ.ОБР</a> | Возвращает наименьшее значение, при котором интегральное биномиальное распределение будет меньше заданного критерия или равно ему |
| <a href="https://support.microsoft.com/office/8a2be3d7-91c3-4b48-9517-64548008563a" target="_blank">Функция БИТ.И</a> | Возвращает результат операции поразрядного И для двух чисел |
| <a href="https://support.microsoft.com/office/c55bb27e-cacd-4c7c-b258-d80861a03c9c" target="_blank">Функция БИТ.СДВИГЛ</a> | Возвращает число со сдвигом влево на указанное число бит |
| <a href="https://support.microsoft.com/office/f6ead5c8-5b98-4c9e-9053-8ad5234919b2" target="_blank">Функция БИТ.ИЛИ</a> | Возвращает результат операции поразрядного ИЛИ для двух чисел |
| <a href="https://support.microsoft.com/office/274d6996-f42c-4743-abdb-4ff95351222c" target="_blank">Функция БИТ.СДВИГП</a> | Возвращает число со сдвигом вправо на указанное число бит |
| <a href="https://support.microsoft.com/office/c81306a1-03f9-4e89-85ac-b86c3cba10e4" target="_blank">Функция БИТ.ИСКЛИЛИ</a> | Возвращает результат операции поразрядного исключающего ИЛИ для двух чисел |
| <a href="https://support.microsoft.com/office/80f95d2f-b499-4eee-9f16-f795a8e306c8" target="_blank">ПОТОЛОК. MATH, ECMA_CEILING функции</a> | Округляет число к большему до ближайшего целого или до ближайшего кратного значения с указанной точностью |
| <a href="https://support.microsoft.com/office/f366a774-527a-4c92-ba49-af0a196e66cb" target="_blank">Функция ОКРВВЕРХ.ТОЧН</a> | Округляет число до ближайшего целого или до ближайшего кратного значения с указанной точностью. Число округляется до большего значения вне зависимости от его знака. |
| <a href="https://support.microsoft.com/office/bbd249c8-b36e-4a91-8017-1c133f9b837a" target="_blank">Функция СИМВОЛ</a> | Возвращает символ с указанным кодом |
| <a href="https://support.microsoft.com/office/8486b05e-5c05-4942-a9ea-f6b341518732" target="_blank">Функция ХИ2.РАСП</a> | Возвращает интегральную функцию плотности бета-распределения |
| <a href="https://support.microsoft.com/office/dc4832e8-ed2b-49ae-8d7c-b28d5804c0f2" target="_blank">Функция ХИ2.РАСП.ПХ</a> | Возвращает одностороннюю вероятность распределения хи-квадрат |
| <a href="https://support.microsoft.com/office/400db556-62b3-472d-80b3-254723e7092f" target="_blank">Функция ХИ2.ОБР</a> | Возвращает интегральную функцию плотности бета-распределения |
| <a href="https://support.microsoft.com/office/435b5ed8-98d5-4da6-823f-293e2cbc94fe" target="_blank">Функция ХИ2.ОБР.ПХ</a> | Возвращает значение, обратное односторонней вероятности распределения хи-квадрат |
| <a href="https://support.microsoft.com/office/fc5c184f-cb62-4ec7-a46e-38653b98f5bc" target="_blank">Функция ВЫБОР</a> | Выбирает значение из списка значений |
| <a href="https://support.microsoft.com/office/26f3d7c5-475f-4a9c-90e5-4b8ba987ba41" target="_blank">Функция ПЕЧСИМВ</a> | Удаляет из текста все непечатаемые символы |
| <a href="https://support.microsoft.com/office/c32b692b-2ed0-4a04-bdd9-75640144b928" target="_blank">Функция КОДСИМВ</a> | Возвращает числовой код первого символа в текстовой строке |
| <a href="https://support.microsoft.com/office/4e8e7b4e-e603-43e8-b177-956088fa48ca" target="_blank">Функция ЧИСЛСТОЛБ</a> | Возвращает количество столбцов в ссылке |
| <a href="https://support.microsoft.com/office/12a3f276-0a21-423a-8de6-06990aaf638a" target="_blank">Функция ЧИСЛКОМБ</a> | Возвращает количество комбинаций, которые можно составить из заданного числа объектов |
| <a href="https://support.microsoft.com/office/efb49eaa-4f4c-4cd2-8179-0ddfcf9d035d" target="_blank">Функция ЧИСЛКОМБА</a> | Возвращает количество комбинаций, которые можно составить из заданного числа элементов, с повторами |
| <a href="https://support.microsoft.com/office/f0b8f3a9-51cc-4d6d-86fb-3a9362fa4128" target="_blank">Функция КОМПЛЕКСН</a> | Преобразует коэффициенты при вещественной и мнимой частях комплексного числа в комплексное число |
| <a href="https://support.microsoft.com/office/8f8ae884-2ca8-4f7a-b093-75d702bea31d" target="_blank">Функция СЦЕПИТЬ</a> | Объединяет несколько текстовых элементов в один |
| <a href="https://support.microsoft.com/office/7cec58a6-85bb-488d-91c3-63828d4fbfd4" target="_blank">Функция ДОВЕРИТ.НОРМ</a> | Возвращает доверительный интервал для среднего генеральной совокупности |
| <a href="https://support.microsoft.com/office/e8eca395-6c3a-4ba9-9003-79ccc61d3c53" target="_blank">Функция ДОВЕРИТ.СТЬЮДЕНТ</a> | Возвращает доверительный интервал для среднего генеральной совокупности, используя распределение Стьюдента |
| <a href="https://support.microsoft.com/office/d785bef1-808e-4aac-bdcd-666c810f9af2" target="_blank">Функция ПРЕОБР</a> | Преобразует значение из одной системы измерения в другую |
| <a href="https://support.microsoft.com/office/0fb808a5-95d6-4553-8148-22aebdce5f05" target="_blank">Функция COS</a> | Возвращает косинус числа |
| <a href="https://support.microsoft.com/office/e460d426-c471-43e8-9540-a57ff3b70555" target="_blank">Функция COSH</a> | Возвращает гиперболический косинус числа |
| <a href="https://support.microsoft.com/office/c446f34d-6fe4-40dc-84f8-cf59e5f5e31a" target="_blank">Функция COT</a> | Возвращает котангенс угла |
| <a href="https://support.microsoft.com/office/2e0b4cb6-0ba0-403e-aed4-deaa71b49df5" target="_blank">Функция COTH</a> | Возвращает гиперболический котангенс числа |
| <a href="https://support.microsoft.com/office/a59cd7fc-b623-4d93-87a4-d23bf411294c" target="_blank">Функция СЧЁТ</a> | Подсчитывает количество чисел в списке аргументов |
| <a href="https://support.microsoft.com/office/7dc98875-d5c1-46f1-9a82-53f3219e2509" target="_blank">Функция СЧЁТЗ</a> | Подсчитывает количество значений в списке аргументов |
| <a href="https://support.microsoft.com/office/6a92d772-675c-4bee-b346-24af6bd3ac22" target="_blank">Функция СЧИТАТЬПУСТОТЫ</a> | Подсчитывает количество пустых ячеек в диапазоне |
| <a href="https://support.microsoft.com/office/e0de10c6-f885-4e71-abb4-1f464816df34" target="_blank">Функция СЧЁТЕСЛИ</a> | Подсчитывает количество ячеек в диапазоне, соответствующих определенному условию |
| <a href="https://support.microsoft.com/office/dda3dc6e-f74e-4aee-88bc-aa8c2a866842" target="_blank">Функция СЧЁТЕСЛИМН</a> | Подсчитывает количество ячеек в диапазоне, соответствующих нескольким условиям |
| <a href="https://support.microsoft.com/office/eb9a8dfb-2fb2-4c61-8e5d-690b320cf872" target="_blank">Функция ДНЕЙКУПОНДО</a> | Возвращает количество дней с начала купонного периода до даты расчета |
| <a href="https://support.microsoft.com/office/cc64380b-315b-4e7b-950c-b30b0a76f671" target="_blank">Функция ДНЕЙКУПОН</a> | Возвращает количество дней расчета в купонном периоде |
| <a href="https://support.microsoft.com/office/5ab3f0b2-029f-4a8b-bb65-47d525eea547" target="_blank">Функция ДНЕЙКУПОНПОСЛЕ</a> | Возвращает количество дней между датой расчета и следующей датой выплаты процентов |
| <a href="https://support.microsoft.com/office/fd962fef-506b-4d9d-8590-16df5393691f" target="_blank">Функция ДАТАКУПОНПОСЛЕ</a> | Возвращает дату выплаты процентов, следующую после даты расчета |
| <a href="https://support.microsoft.com/office/a90af57b-de53-4969-9c99-dd6139db2522" target="_blank">Функция ЧИСЛКУПОН</a> | Возвращает количество процентных выплат между датой расчета и датой погашения |
| <a href="https://support.microsoft.com/office/2eb50473-6ee9-4052-a206-77a9a385d5b3" target="_blank">Функция ДАТАКУПОНДО</a> | Возвращает дату выплаты процентов, которая предшествует дате расчета |
| <a href="https://support.microsoft.com/office/07379361-219a-4398-8675-07ddc4f135c1" target="_blank">Функция CSC</a> | Возвращает косеканс угла |
| <a href="https://support.microsoft.com/office/f58f2c22-eb75-4dd6-84f4-a503527f8eeb" target="_blank">Функция CSCH</a> | Возвращает гиперболический косеканс угла |
| <a href="https://support.microsoft.com/office/61067bb0-9016-427d-b95b-1a752af0e606" target="_blank">Функция ОБЩПЛАТ</a> | Возвращает кумулятивную сумму процентов, выплачиваемую между двумя периодами |
| <a href="https://support.microsoft.com/office/94a4516d-bd65-41a1-bc16-053a6af4c04d" target="_blank">Функция ОБЩДОХОД</a> | Возвращает кумулятивную сумму, выплачиваемую для погашения займа между двумя периодами |
| <a href="https://support.microsoft.com/office/e36c0c8c-4104-49da-ab83-82328b832349" target="_blank">Функция ДАТА</a> | Возвращает порядковый номер определенной даты |
| <a href="https://support.microsoft.com/office/df8b07d4-7761-4a93-bc33-b7471bbff252" target="_blank">Функция ДАТАЗНАЧ</a> | Преобразует дату из текстового формата в числовой |
| <a href="https://support.microsoft.com/office/a6a2d5ac-4b4b-48cd-a1d8-7b37834e5aee" target="_blank">Функция ДСРЗНАЧ</a> | Возвращает среднее значение выбранных записей базы данных |
| <a href="https://support.microsoft.com/office/8a7d1cbb-6c7d-4ba1-8aea-25c134d03101" target="_blank">Функция ДЕНЬ</a> | Преобразует порядковый номер в день месяца |
| <a href="https://support.microsoft.com/office/57740535-d549-4395-8728-0f07bff0b9df" target="_blank">Функция ДНИ</a> | Возвращает количество дней между двумя датами |
| <a href="https://support.microsoft.com/office/b9a509fd-49ef-407e-94df-0cbda5718c2a" target="_blank">Функция ДНЕЙ360</a> | Вычисляет количество дней между двумя датами на основании 360-дневного года |
| <a href="https://support.microsoft.com/office/354e7d28-5f93-4ff1-8a52-eb4ee549d9d7" target="_blank">Функция ФУО</a> | Возвращает сумму амортизации актива за определенный период, начисляемую по методу фиксированного убывающего остатка |
| <a href="https://support.microsoft.com/office/a4025e73-63d2-4958-9423-21a24794c9e5" target="_blank">Функция DBCS</a> | Преобразует полуширинные (однобайтовые) английские буквы или знаки катакана в пределах строки символов в полноширинные (двухбайтовые) символы |
| <a href="https://support.microsoft.com/office/c1fc7b93-fb0d-4d8d-97db-8d5f076eaeb1" target="_blank">Функция БСЧЁТ</a> | Подсчитывает количество ячеек в базе данных, содержащих числа |
| <a href="https://support.microsoft.com/office/00232a6d-5a66-4a01-a25b-c1653fda1244" target="_blank">Функция БСЧЁТА</a> | Подсчитывает количество непустых ячеек в базе данных |
| <a href="https://support.microsoft.com/office/519a7a37-8772-4c96-85c0-ed2c209717a5" target="_blank">Функция ДДОБ</a> | Возвращает сумму амортизации актива за определенный период, начисляемую методом двойного убывающего остатка или иным указанным методом |
| <a href="https://support.microsoft.com/office/0f63dd0e-5d1a-42d8-b511-5bf5c6d43838" target="_blank">Функция ДЕС.В.ДВ</a> | Преобразует десятичное число в двоичное |
| <a href="https://support.microsoft.com/office/6344ee8b-b6b5-4c6a-a672-f64666704619" target="_blank">Функция ДЕС.В.ШЕСТН</a> | Преобразует десятичное число в шестнадцатеричное |
| <a href="https://support.microsoft.com/office/c9d835ca-20b7-40c4-8a9e-d3be351ce00f" target="_blank">Функция ДЕС.В.ВОСЬМ</a> | Преобразует десятичное число в восьмеричное |
| <a href="https://support.microsoft.com/office/ee554665-6176-46ef-82de-0a283658da2e" target="_blank">Функция ДЕС</a> | Преобразует текстовое представление числа c указанным основанием в десятичное |
| <a href="https://support.microsoft.com/office/4d6ec4db-e694-4b94-ace0-1cc3f61f9ba1" target="_blank">Функция ГРАДУСЫ</a> | Преобразует радианы в градусы |
| <a href="https://support.microsoft.com/office/2f763672-c959-4e07-ac33-fe03220ba432" target="_blank">Функция ДЕЛЬТА</a> | Проверяет равенство двух значений |
| <a href="https://support.microsoft.com/office/8b739616-8376-4df5-8bd0-cfe0a6caf444" target="_blank">Функция КВАДРОТКЛ</a> | Возвращает сумму квадратов отклонений |
| <a href="https://support.microsoft.com/office/455568bf-4eef-45f7-90f0-ec250d00892e" target="_blank">Функция БИЗВЛЕЧЬ</a> | Извлекает из базы данных одну запись, соответствующую заданному условию |
| <a href="https://support.microsoft.com/office/71fce9f3-3f05-4acf-a5a3-eac6ef4daa53" target="_blank">Функция СКИДКА</a> | Возвращает ставку дисконтирования ценной бумаги |
| <a href="https://support.microsoft.com/office/f4e8209d-8958-4c3d-a1ee-6351665d41c2" target="_blank">Функция ДМАКС</a> | Возвращает наибольшее значение из выбранных записей базы данных |
| <a href="https://support.microsoft.com/office/4ae6f1d9-1f26-40f1-a783-6dc3680192a3" target="_blank">Функция ДМИН</a> | Возвращает наименьшее значение из выбранных записей базы данных |
| <a href="https://support.microsoft.com/office/a6cd05d9-9740-4ad3-a469-8109d18ff611" target="_blank">Функции DOLLAR, USDOLLAR</a> | Преобразует число в текст, используя денежный формат |
| <a href="https://support.microsoft.com/office/db85aab0-1677-428a-9dfd-a38476693427" target="_blank">Функция РУБЛЬ.ДЕС</a> | Преобразует цену в рублях, представленную в виде десятичной дроби, в десятичное число |
| <a href="https://support.microsoft.com/office/0835d163-3023-4a33-9824-3042c5d4f495" target="_blank">Функция РУБЛЬ.ДРОБЬ</a> | Преобразует цену в рублях, представленную в виде десятичного числа, в десятичную дробь |
| <a href="https://support.microsoft.com/office/4f96b13e-d49c-47a7-b769-22f6d017cb31" target="_blank">Функция БДПРОИЗВЕД</a> | Перемножает значения определенного поля записей, соответствующих условию, в базе данных |
| <a href="https://support.microsoft.com/office/026b8c73-616d-4b5e-b072-241871c4ab96" target="_blank">Функция ДСТАНДОТКЛ</a> | Оценивает стандартное отклонение для выборки записей базы данных |
| <a href="https://support.microsoft.com/office/04b78995-da03-4813-bbd9-d74fd0f5d94b" target="_blank">Функция ДСТАНДОТКЛП</a> | Вычисляет стандартное отклонение для генеральной совокупности выбранных записей базы данных |
| <a href="https://support.microsoft.com/office/53181285-0c4b-4f5a-aaa3-529a322be41b" target="_blank">Функция БДСУММ</a> | Суммирует числа в поле (столбце) записей базы данных, соответствующих условию |
| <a href="https://support.microsoft.com/office/b254ea57-eadc-4602-a86a-c8e369334038" target="_blank">Функция ДЛИТ</a> | Возвращает дюрацию ценной бумаги с периодической выплатой процентов в годовом исчислении |
| <a href="https://support.microsoft.com/office/d6747ca9-99c7-48bb-996e-9d7af00f3ed1" target="_blank">Функция БДДИСП</a> | Оценивает дисперсию для выборки записей базы данных |
| <a href="https://support.microsoft.com/office/eb0ba387-9cb7-45c8-81e9-0394912502fc" target="_blank">Функция БДДИСПП</a> | Вычисляет дисперсию для генеральной совокупности выбранных записей базы данных |
| <a href="https://support.microsoft.com/office/3c920eb2-6e66-44e7-a1f5-753ae47ee4f5" target="_blank">Функция ДАТАМЕС</a> | Возвращает порядковый номер даты, отстоящей на заданное количество месяцев вперед или назад от начальной даты |
| <a href="https://support.microsoft.com/office/910d4e4c-79e2-4009-95e6-507e04f11bc4" target="_blank">Функция ЭФФЕКТ</a> | Возвращает эффективную годовую процентную ставку |
| <a href="https://support.microsoft.com/office/7314ffa1-2bc9-4005-9d66-f49db127d628" target="_blank">Функция КОНМЕСЯЦА</a> | Возвращает порядковый номер последнего дня месяца, отстоящего на заданное число месяцев вперед или назад от начальной даты |
| <a href="https://support.microsoft.com/office/c53c7e7b-5482-4b6c-883e-56df3c9af349" target="_blank">Функция ФОШ</a> | Возвращает функцию ошибок |
| <a href="https://support.microsoft.com/office/9a349593-705c-4278-9a98-e4122831a8e0" target="_blank">Функция ФОШ.ТОЧН</a> | Возвращает функцию ошибок |
| <a href="https://support.microsoft.com/office/736e0318-70ba-4e8b-8d08-461fe68b71b3" target="_blank">Функция ДФОШ</a> | Возвращает дополнительную функцию ошибок |
| <a href="https://support.microsoft.com/office/e90e6bab-f45e-45df-b2ac-cd2eb4d4a273" target="_blank">Функция ДФОШ.ТОЧН</a> | Возвращает дополнительную функцию ошибок, проинтегрированную от x до бесконечности |
| <a href="https://support.microsoft.com/office/10958677-7c8d-44f7-ae77-b9a9ee6eefaa" target="_blank">Функция ТИП.ОШИБКИ</a> | Возвращает номер, соответствующий типу ошибки |
| <a href="https://support.microsoft.com/office/197b5f06-c795-4c1e-8696-3c3b8a646cf9" target="_blank">Функция ЧЁТН</a> | Округляет число к большему до ближайшего четного целого |
| <a href="https://support.microsoft.com/office/d3087698-fc15-4a15-9631-12575cf29926" target="_blank">Функция СОВПАД</a> | Проверяет идентичность двух текстовых значений |
| <a href="https://support.microsoft.com/office/c578f034-2c45-4c37-bc8c-329660a63abe" target="_blank">Функция EXP</a> | Возвращает число e, возведенное в указанную степень |
| <a href="https://support.microsoft.com/office/4c12ae24-e563-4155-bf3e-8b78b6ae140e" target="_blank">Функция ЭКСП.РАСП</a> | Возвращает экспоненциальное распределение |
| <a href="https://support.microsoft.com/office/a887efdc-7c8e-46cb-a74a-f884cd29b25d" target="_blank">Функция F.РАСП</a> | Возвращает F-распределение вероятности |
| <a href="https://support.microsoft.com/office/d74cbb00-6017-4ac9-b7d7-6049badc0520" target="_blank">Функция F.РАСП.ПХ</a> | Возвращает F-распределение вероятности |
| <a href="https://support.microsoft.com/office/0dda0cf9-4ea0-42fd-8c3c-417a1ff30dbe" target="_blank">Функция F.ОБР</a> | Возвращает обратное F-распределение вероятности |
| <a href="https://support.microsoft.com/office/d371aa8f-b0b1-40ef-9cc2-496f0693ac00" target="_blank">Функция F.ОБР.ПХ</a> | Возвращает обратное F-распределение вероятности |
| <a href="https://support.microsoft.com/office/ca8588c2-15f2-41c0-8e8c-c11bd471a4f3" target="_blank">Функция ФАКТР</a> | Возвращает факториал числа |
| <a href="https://support.microsoft.com/office/e67697ac-d214-48eb-b7b7-cce2589ecac8" target="_blank">Функция ДВФАКТР</a> | Возвращает двойной факториал числа |
| <a href="https://support.microsoft.com/office/2d58dfa5-9c03-4259-bf8f-f0ae14346904" target="_blank">Функция ЛОЖЬ</a> | Возвращает логическое значение `FALSE` |
| <a href="https://support.microsoft.com/office/c7912941-af2a-4bdf-a553-d0d89b0a0628" target="_blank">Функции НАЙТИ, НАЙТИБ</a> | Находит одно текстовое значение в другом (с учетом регистра) |
| <a href="https://support.microsoft.com/office/d656523c-5076-4f95-b87b-7741bf236c69" target="_blank">Функция ФИШЕР</a> | Возвращает преобразование Фишера |
| <a href="https://support.microsoft.com/office/62504b39-415a-4284-a285-19c8e82f86bb" target="_blank">Функция ФИШЕРОБР</a> | Возвращает обратное преобразование Фишера |
| <a href="https://support.microsoft.com/office/ffd5723c-324c-45e9-8b96-e41be2a8274a" target="_blank">Функция ФИКСИРОВАННЫЙ</a> | Форматирует число, отображая определенное количество знаков после запятой |
| <a href="https://support.microsoft.com/office/c302b599-fbdb-4177-ba19-2c2b1249a2f5" target="_blank">Функция ОКРВНИЗ.МАТ</a> | Округляет число к меньшему до ближайшего целого или до ближайшего кратного с указанной точностью |
| <a href="https://support.microsoft.com/office/f769b468-1452-4617-8dc3-02f842a0702e" target="_blank">Функция ОКРВНИЗ.ТОЧН</a> | Округляет число к меньшему до ближайшего целого или до ближайшего кратного с указанной точностью. Число округляется до меньшего значения вне зависимости от его знака. |
| <a href="https://support.microsoft.com/office/2eef9f44-a084-4c61-bdd8-4fe4bb1b71b3" target="_blank">Функция БС</a> | Возвращает будущую стоимость инвестиций |
| <a href="https://support.microsoft.com/office/bec29522-bd87-4082-bab9-a241f3fb251d" target="_blank">Функция БЗРАСПИС</a> | Возвращает будущую стоимость первоначальной основной суммы после применения ряда ставок сложных процентов |
| <a href="https://support.microsoft.com/office/ce1702b1-cf55-471d-8307-f83be0fc5297" target="_blank">Функция ГАММА</a> | Возвращает значение гамма-функции |
| <a href="https://support.microsoft.com/office/9b6f1538-d11c-4d5f-8966-21f6a2201def" target="_blank">Функция ГАММА.РАСП</a> | Возвращает гамма-распределение |
| <a href="https://support.microsoft.com/office/74991443-c2b0-4be5-aaab-1aa4d71fbb18" target="_blank">Функция ГАММА.ОБР</a> | Возвращает обратное интегральное гамма-распределение |
| <a href="https://support.microsoft.com/office/b838c48b-c65f-484f-9e1d-141c55470eb9" target="_blank">Функция ГАММАНЛОГ</a> | Возвращает натуральный логарифм гамма-функции — Γ(x) |
| <a href="https://support.microsoft.com/office/5cdfe601-4e1e-4189-9d74-241ef1caa599" target="_blank">Функция ГАММАНЛОГ.ТОЧН</a> | Возвращает натуральный логарифм гамма-функции — Γ(x) |
| <a href="https://support.microsoft.com/office/069f1b4e-7dee-4d6a-a71f-4b69044a6b33" target="_blank">Функция ГАУСС</a> | Возвращает значение на 0,5 меньше стандартного нормального интегрального распределения |
| <a href="https://support.microsoft.com/office/d5107a51-69e3-461f-8e4c-ddfc21b5073a" target="_blank">Функция НОД</a> | Возвращает наибольший общий делитель |
| <a href="https://support.microsoft.com/office/db1ac48d-25a5-40a0-ab83-0b38980e40d5" target="_blank">Функция СРГЕОМ</a> | Возвращает среднее геометрическое значение |
| <a href="https://support.microsoft.com/office/f37e7d2a-41da-4129-be95-640883fca9df" target="_blank">Функция ПОРОГ</a> | Проверяет, превышает ли число пороговое значение |
| <a href="https://support.microsoft.com/office/5efd9184-fab5-42f9-b1d3-57883a1d3bc6" target="_blank">Функция СРГАРМ</a> | Возвращает среднее гармоническое значение |
| <a href="https://support.microsoft.com/office/a13aafaa-5737-4920-8424-643e581828c1" target="_blank">Функция ШЕСТН.В.ДВ</a> | Преобразует шестнадцатеричное число в двоичное |
| <a href="https://support.microsoft.com/office/8c8c3155-9f37-45a5-a3ee-ee5379ef106e" target="_blank">Функция ШЕСТН.В.ДЕС</a> | Преобразует шестнадцатеричное число в десятичное |
| <a href="https://support.microsoft.com/office/54d52808-5d19-4bd0-8a63-1096a5d11912" target="_blank">Функция ШЕСТН.В.ВОСЬМ</a> | Преобразует шестнадцатеричное число в восьмеричное |
| <a href="https://support.microsoft.com/office/a3034eec-b719-4ba3-bb65-e1ad662ed95f" target="_blank">Функция ГПР</a> | Выполняет поиск в первой строке массива и возвращает значение указанной ячейки |
| <a href="https://support.microsoft.com/office/a3afa879-86cb-4339-b1b5-2dd2d7310ac7" target="_blank">Функция ЧАС</a> | Преобразует порядковый номер в час |
| <a href="https://support.microsoft.com/office/333c7ce6-c5ae-4164-9c47-7de9b76f577f" target="_blank">Функция ГИПЕРССЫЛКА</a> | Создает гиперссылку на документ, расположенный на сетевом сервере, в интрасети или Интернете |
| <a href="https://support.microsoft.com/office/6dbd547f-1d12-4b1f-8ae5-b0d9e3d22fbf" target="_blank">Функция ГИПЕРГЕОМ.РАСП</a> | Возвращает гипергеометрическое распределение |
| <a href="https://support.microsoft.com/office/69aed7c9-4e8a-4755-a9bc-aa8bbff73be2" target="_blank">Функция ЕСЛИ</a> | Выполняет указанную логическую проверку |
| <a href="https://support.microsoft.com/office/b31e73c6-d90c-4062-90bc-8eb351d765a1" target="_blank">Функция МНИМ.ABS</a> | Возвращает абсолютную величину (модуль) комплексного числа |
| <a href="https://support.microsoft.com/office/dd5952fd-473d-44d9-95a1-9a17b23e428a" target="_blank">Функция МНИМ.ЧАСТЬ</a> | Возвращает коэффициент при мнимой части комплексного числа |
| <a href="https://support.microsoft.com/office/eed37ec1-23b3-4f59-b9f3-d340358a034a" target="_blank">Функция МНИМ.АРГУМЕНТ</a> | Возвращает аргумент тета — угол, выраженный в радианах |
| <a href="https://support.microsoft.com/office/2e2fc1ea-f32b-4f9b-9de6-233853bafd42" target="_blank">Функция МНИМ.СОПРЯЖ</a> | Возвращает комплексно-сопряженное число для комплексного числа |
| <a href="https://support.microsoft.com/office/dad75277-f592-4a6b-ad6c-be93a808a53c" target="_blank">Функция МНИМ.COS</a> | Возвращает косинус комплексного числа |
| <a href="https://support.microsoft.com/office/053e4ddb-4122-458b-be9a-457c405e90ff" target="_blank">Функция МНИМ.COSH</a> | Возвращает гиперболический косинус комплексного числа |
| <a href="https://support.microsoft.com/office/dc6a3607-d26a-4d06-8b41-8931da36442c" target="_blank">Функция МНИМ.COT</a> | Возвращает котангенс комплексного числа |
| <a href="https://support.microsoft.com/office/9e158d8f-2ddf-46cd-9b1d-98e29904a323" target="_blank">Функция МНИМ.CSC</a> | Возвращает косеканс комплексного числа |
| <a href="https://support.microsoft.com/office/c0ae4f54-5f09-4fef-8da0-dc33ea2c5ca9" target="_blank">Функция МНИМ.CSCH</a> | Возвращает гиперболический косеканс комплексного числа |
| <a href="https://support.microsoft.com/office/a505aff7-af8a-4451-8142-77ec3d74d83f" target="_blank">Функция МНИМ.ДЕЛ</a> | Возвращает частное от деления двух комплексных чисел |
| <a href="https://support.microsoft.com/office/c6f8da1f-e024-4c0c-b802-a60e7147a95f" target="_blank">Функция МНИМ.EXP</a> | Возвращает экспоненту комплексного числа |
| <a href="https://support.microsoft.com/office/32b98bcf-8b81-437c-a636-6fb3aad509d8" target="_blank">Функция МНИМ.LN</a> | Возвращает натуральный логарифм комплексного числа |
| <a href="https://support.microsoft.com/office/58200fca-e2a2-4271-8a98-ccd4360213a5" target="_blank">Функция МНИМ.LOG10</a> | Возвращает десятичный логарифм комплексного числа |
| <a href="https://support.microsoft.com/office/152e13b4-bc79-486c-a243-e6a676878c51" target="_blank">Функция МНИМ.LOG2</a> | Возвращает двоичный логарифм комплексного числа |
| <a href="https://support.microsoft.com/office/210fd2f5-f8ff-4c6a-9d60-30e34fbdef39" target="_blank">Функция МНИМ.СТЕПЕНЬ</a> | Возвращает комплексное число, возведенное в степень с целочисленным показателем |
| <a href="https://support.microsoft.com/office/2fb8651a-a4f2-444f-975e-8ba7aab3a5ba" target="_blank">Функция МНИМ.ПРОИЗВЕД</a> | Возвращает произведение от 2 до 255 комплексных чисел |
| <a href="https://support.microsoft.com/office/d12bc4c0-25d0-4bb3-a25f-ece1938bf366" target="_blank">Функция МНИМ.ВЕЩ</a> | Возвращает коэффициент при вещественной части комплексного числа |
| <a href="https://support.microsoft.com/office/6df11132-4411-4df4-a3dc-1f17372459e0" target="_blank">Функция МНИМ.SEC</a> | Возвращает секанс комплексного числа |
| <a href="https://support.microsoft.com/office/f250304f-788b-4505-954e-eb01fa50903b" target="_blank">Функция МНИМ.SECH</a> | Возвращает гиперболический секанс комплексного числа |
| <a href="https://support.microsoft.com/office/1ab02a39-a721-48de-82ef-f52bf37859f6" target="_blank">Функция МНИМ.SIN</a> | Возвращает синус комплексного числа |
| <a href="https://support.microsoft.com/office/dfb9ec9e-8783-4985-8c42-b028e9e8da3d" target="_blank">Функция МНИМ.SINH</a> | Возвращает гиперболический синус комплексного числа |
| <a href="https://support.microsoft.com/office/e1753f80-ba11-4664-a10e-e17368396b70" target="_blank">Функция МНИМ.КОРЕНЬ</a> | Возвращает значение квадратного корня из комплексного числа |
| <a href="https://support.microsoft.com/office/2e404b4d-4935-4e85-9f52-cb08b9a45054" target="_blank">Функция МНИМ.РАЗН</a> | Возвращает разность двух комплексных чисел |
| <a href="https://support.microsoft.com/office/81542999-5f1c-4da6-9ffe-f1d7aaa9457f" target="_blank">Функция МНИМ.СУММ</a> | Возвращает сумму комплексных чисел |
| <a href="https://support.microsoft.com/office/8478f45d-610a-43cf-8544-9fc0b553a132" target="_blank">Функция МНИМ.TAN</a> | Возвращает тангенс комплексного числа |
| <a href="https://support.microsoft.com/office/a6c4af9e-356d-4369-ab6a-cb1fd9d343ef" target="_blank">Функция ЦЕЛОЕ</a> | Округляет число к меньшему до ближайшего целого |
| <a href="https://support.microsoft.com/office/5cb34dde-a221-4cb6-b3eb-0b9e55e1316f" target="_blank">Функция ИНОРМА</a> | Возвращает процентную ставку для полностью инвестированной ценной бумаги |
| <a href="https://support.microsoft.com/office/5cce0ad6-8402-4a41-8d29-61a0b054cb6f" target="_blank">Функция ПРПЛТ</a> | Возвращает сумму процентных выплат по инвестиции за определенный период |
| <a href="https://support.microsoft.com/office/64925eaa-9988-495b-b290-3ad0c163c1bc" target="_blank">Функция ВСД</a> | Возвращает внутреннюю норму доходности на основании ряда денежных потоков. |
| <a href="https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Функция ЕОШ</a> | Возвращает значение `TRUE`, если ячейка содержит ошибку (кроме #Н/Д) |
| <a href="https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Функция ЕОШИБКА</a> | Возвращает значение `TRUE`, если ячейка содержит ошибку |
| <a href="https://support.microsoft.com/office/aa15929a-d77b-4fbb-92f4-2f479af55356" target="_blank">Функция ЕЧЁТН</a> | Возвращает значение `TRUE`, если ячейка содержит четное число |
| <a href="https://support.microsoft.com/office/e4d1355f-7121-4ef2-801e-3839bfd6b1e5" target="_blank">Функция ЕФОРМУЛА</a> | Возвращает значение `TRUE`, если ячейка содержит формулу |
| <a href="https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Функция ЕЛОГИЧ</a> | Возвращает значение `TRUE`, если ячейка содержит логическое значение |
| <a href="https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Функция ЕНД</a> | Возвращает значение `TRUE`, если ячейка содержит ошибку #Н/Д |
| <a href="https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Функция ЕНЕТЕКСТ</a> | Возвращает значение `TRUE`, если ячейка содержит любое значение, кроме текстового |
| <a href="https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Функция ЕЧИСЛО</a> | Возвращает значение `TRUE`, если ячейка содержит число |
| <a href="https://support.microsoft.com/office/e587bb73-6cc2-4113-b664-ff5b09859a83" target="_blank">Функция ISO.ОКРВВЕРХ</a> | Возвращает число, округленное к большему до ближайшего целого или до ближайшего кратного с указанной точностью |
| <a href="https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Функция ЕНЕЧЁТ</a> | Возвращает значение `TRUE`, если ячейка содержит нечетное число |
| <a href="https://support.microsoft.com/office/1c2d0afe-d25b-4ab1-8894-8d0520e90e0e" target="_blank">Функция НОМНЕДЕЛИ.ISO</a> | Возвращает номер недели в году для определенной даты в соответствии со стандартами ISO |
| <a href="https://support.microsoft.com/office/fa58adb6-9d39-4ce0-8f43-75399cea56cc" target="_blank">Функция ПРОЦПЛАТ</a> | Вычисляет сумму процентов, выплачиваемую в течение определенного инвестиционного периода |
| <a href="https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Функция ЕССЫЛКА</a> | Возвращает значение `TRUE`, если ячейка содержит ссылку |
| <a href="https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Функция ЕТЕКСТ</a> | Возвращает значение `TRUE`, если ячейка содержит текстовое значение |
| <a href="https://support.microsoft.com/office/bc3a265c-5da4-4dcb-b7fd-c237789095ab" target="_blank">Функция ЭКСЦЕСС</a> | Возвращает эксцесс множества данных |
| <a href="https://support.microsoft.com/office/3af0af19-1190-42bb-bb8b-01672ec00a64" target="_blank">Функция НАИБОЛЬШИЙ</a> | Возвращает k-е наибольшее значение в наборе данных |
| <a href="https://support.microsoft.com/office/7152b67a-8bb5-4075-ae5c-06ede5563c94" target="_blank">Функция НОК</a> | Возвращает наименьшее общее кратное |
| <a href="https://support.microsoft.com/office/9203d2d2-7960-479b-84c6-1ea52b99640c" target="_blank">Функции ЛЕВСИМВ, ЛЕВБ</a> | Возвращают первые символы в текстовой строке |
| <a href="https://support.microsoft.com/office/29236f94-cedc-429d-affd-b5e33d2c67cb" target="_blank">Функции ДЛСТР, ДЛИНБ</a> | Возвращает количество символов в текстовой строке |
| <a href="https://support.microsoft.com/office/81fe1ed7-dac9-4acd-ba1d-07a142c6118f" target="_blank">Функция LN</a> | Возвращает натуральный логарифм числа |
| <a href="https://support.microsoft.com/office/4e82f196-1ca9-4747-8fb0-6c4a3abb3280" target="_blank">Функция LOG</a> | Возвращает логарифм числа по заданному основанию |
| <a href="https://support.microsoft.com/office/c75b881b-49dd-44fb-b6f4-37e3486a0211" target="_blank">Функция LOG10</a> | Возвращает десятичный логарифм числа |
| <a href="https://support.microsoft.com/office/eb60d00b-48a9-4217-be2b-6074aee6b070" target="_blank">Функция ЛОГНОРМ.РАСП</a> | Возвращает интегральное логнормальное распределение |
| <a href="https://support.microsoft.com/office/fe79751a-f1f2-4af8-a0a1-e151b2d4f600" target="_blank">Функция ЛОГНОРМ.ОБР</a> | Возвращает обратное интегральное логнормальное распределение |
| <a href="https://support.microsoft.com/office/446d94af-663b-451d-8251-369d5e3864cb" target="_blank">Функция ПРОСМОТР</a> | Ищет значения в строке, столбце или массиве |
| <a href="https://support.microsoft.com/office/3f21df02-a80c-44b2-afaf-81358f9fdeb4" target="_blank">Функция СТРОЧН</a> | Преобразует текст в нижний регистр |
| <a href="https://support.microsoft.com/office/e8dffd45-c762-47d6-bf89-533f4a37673a" target="_blank">Функция ПОИСКПОЗ</a> | Ищет значения в ссылке или массиве |
| <a href="https://support.microsoft.com/office/e0012414-9ac8-4b34-9a47-73e662c08098" target="_blank">Функция МАКС</a> | Возвращает максимальное значение в списке аргументов |
| <a href="https://support.microsoft.com/office/814bda1e-3840-4bff-9365-2f59ac2ee62d" target="_blank">Функция МАКСА</a> | Возвращает максимальное значение в списке аргументов (включая числовые, текстовые и логические) |
| <a href="https://support.microsoft.com/office/b3786a69-4f20-469a-94ad-33e5b90a763c" target="_blank">Функция МДЛИТ</a> | Возвращает модифицированную дюрацию Маколея для ценной бумаги с предполагаемой номинальной стоимостью 100 р. |
| <a href="https://support.microsoft.com/office/d0916313-4753-414c-8537-ce85bdd967d2" target="_blank">Функция МЕДИАНА</a> | Возвращает медиану заданных чисел |
| <a href="https://support.microsoft.com/office/d5f9e25c-d7d6-472e-b568-4ecb12433028" target="_blank">Функции ПСТР, ПСТРБ</a> | Возвращают определенное количество знаков из текстовой строки, начиная с указанной позиции |
| <a href="https://support.microsoft.com/office/61635d12-920f-4ce2-a70f-96f202dcc152" target="_blank">Функция МИН</a> | Возвращает минимальное значение в списке аргументов |
| <a href="https://support.microsoft.com/office/245a6f46-7ca5-4dc7-ab49-805341bc31d3" target="_blank">Функция МИНА</a> | Возвращает минимальное значение в списке аргументов (включая числовые, текстовые и логические) |
| <a href="https://support.microsoft.com/office/af728df0-05c4-4b07-9eed-a84801a60589" target="_blank">Функция МИНУТЫ</a> | Преобразует порядковый номер в минуты |
| <a href="https://support.microsoft.com/office/b020f038-7492-4fb4-93c1-35c345b53524" target="_blank">Функция МВСД</a> | Возвращает внутреннюю норму доходности с учетом разных ставок финансирования для положительного и отрицательного денежных потоков |
| <a href="https://support.microsoft.com/office/9b6cd169-b6ee-406a-a97b-edf2a9dc24f3" target="_blank">Функция ОСТАТ</a> | Возвращает остаток от деления |
| <a href="https://support.microsoft.com/office/579a2881-199b-48b2-ab90-ddba0eba86e8" target="_blank">Функция МЕСЯЦ</a> | Преобразует порядковый номер в месяц |
| <a href="https://support.microsoft.com/office/c299c3b0-15a5-426d-aa4b-d2d5b3baf427" target="_blank">Функция ОКРУГЛТ</a> | Возвращает число, округленное с заданной точностью |
| <a href="https://support.microsoft.com/office/6fa6373c-6533-41a2-a45e-a56db1db1bf6" target="_blank">Функция МУЛЬТИНОМ</a> | Возвращает полиномиальный коэффициент набора чисел |
| <a href="https://support.microsoft.com/office/a624cad1-3635-4208-b54a-29733d1278c9" target="_blank">Функция Ч</a> | Возвращает значение, преобразованное в число |
| <a href="https://support.microsoft.com/office/5469c2d1-a90c-4fb5-9bbc-64bd9bb6b47c" target="_blank">Функция НД</a> | Возвращает ошибку #Н/Д |
| <a href="https://support.microsoft.com/office/c8239f89-c2d0-45bd-b6af-172e570f8599" target="_blank">Функция ОТРБИНОМ.РАСП</a> | Возвращает отрицательное биномиальное распределение |
| <a href="https://support.microsoft.com/office/48e717bf-a7a3-495f-969e-5005e3eb18e7" target="_blank">Функция ЧИСТРАБДНИ</a> | Возвращает количество полных рабочих дней между двумя датами |
| <a href="https://support.microsoft.com/office/a9b26239-4f20-46a1-9ab8-4e925bfd5e28" target="_blank">Функция ЧИСТРАБДНИ.МЕЖД</a> | Возвращает количество полных рабочих дней между двумя датами с использованием параметров, определяющих, сколько в неделе выходных и какие дни являются выходными |
| <a href="https://support.microsoft.com/office/7f1ae29b-6b92-435e-b950-ad8b190ddd2b" target="_blank">Функция НОМИНАЛ</a> | Возвращает номинальную годовую процентную ставку |
| <a href="https://support.microsoft.com/office/edb1cc14-a21c-4e53-839d-8082074c9f8d" target="_blank">Функция НОРМ.РАСП</a> | Возвращает нормальное распределение |
| <a href="https://support.microsoft.com/office/54b30935-fee7-493c-bedb-2278a9db7e13" target="_blank">Функция НОРМ.ОБР</a> | Возвращает обратное нормальное распределение |
| <a href="https://support.microsoft.com/office/1e787282-3832-4520-a9ae-bd2a8d99ba88" target="_blank">Функция НОРМ.СТ.РАСП</a> | Возвращает стандартное нормальное распределение |
| <a href="https://support.microsoft.com/office/d6d556b4-ab7f-49cd-b526-5a20918452b1" target="_blank">Функция НОРМ.СТ.ОБР</a> | Возвращает обратное стандартное нормальное распределение |
| <a href="https://support.microsoft.com/office/9cfc6011-a054-40c7-a140-cd4ba2d87d77" target="_blank">Функция НЕ</a> | Меняет значение аргумента на противоположное |
| <a href="https://support.microsoft.com/office/3337fd29-145a-4347-b2e6-20c904739c46" target="_blank">Функция ТДАТА</a> | Возвращает порядковый номер текущей даты и времени |
| <a href="https://support.microsoft.com/office/240535b5-6653-4d2d-bfcf-b6a38151d815" target="_blank">Функция КПЕР</a> | Возвращает количество периодов для инвестиций |
| <a href="https://support.microsoft.com/office/8672cb67-2576-4d07-b67b-ac28acf2a568" target="_blank">Функция ЧПС</a> | Возвращает чистую приведенную стоимость инвестиций. Вычисления основываются на ряде периодических денежных потоков и ставки дисконтирования |
| <a href="https://support.microsoft.com/office/1b05c8cf-2bfa-4437-af70-596c7ea7d879" target="_blank">Функция ЧЗНАЧ</a> | Преобразует текст в число без учета языкового стандарта |
| <a href="https://support.microsoft.com/office/55383471-3c56-4d27-9522-1a8ec646c589" target="_blank">Функция ВОСЬМ.В.ДВ</a> | Преобразует восьмеричное число в двоичное |
| <a href="https://support.microsoft.com/office/87606014-cb98-44b2-8dbb-e48f8ced1554" target="_blank">Функция ВОСЬМ.В.ДЕС</a> | Преобразует восьмеричное число в десятичное |
| <a href="https://support.microsoft.com/office/912175b4-d497-41b4-a029-221f051b858f" target="_blank">Функция ВОСЬМ.В.ШЕСТН</a> | Преобразует восьмеричное число в шестнадцатеричное |
| <a href="https://support.microsoft.com/office/deae64eb-e08a-4c88-8b40-6d0b42575c98" target="_blank">Функция НЕЧЁТ</a> | Округляет число к большему до ближайшего нечетного целого |
| <a href="https://support.microsoft.com/office/d7d664a8-34df-4233-8d2b-922bcf6a69e1" target="_blank">Функция ЦЕНАПЕРВНЕРЕГ</a> | Возвращает стоимость ценной бумаги номиналом 100 рублей с нерегулярным первым периодом |
| <a href="https://support.microsoft.com/office/66bc8b7b-6501-4c93-9ce3-2fd16220fe37" target="_blank">Функция ДОХОДПЕРВНЕРЕГ</a> | Возвращает доходность ценной бумаги с нерегулярным первым периодом |
| <a href="https://support.microsoft.com/office/fb657749-d200-4902-afaf-ed5445027fc4" target="_blank">Функция ЦЕНАПОСЛНЕРЕГ</a> | Возвращает стоимость ценной бумаги номиналом 100 рублей с нерегулярным последним периодом |
| <a href="https://support.microsoft.com/office/c873d088-cf40-435f-8d41-c8232fee9238" target="_blank">Функция ДОХОДПОСЛНЕРЕГ</a> | Возвращает доходность ценной бумаги с нерегулярным последним периодом |
| <a href="https://support.microsoft.com/office/7d17ad14-8700-4281-b308-00b131e22af0" target="_blank">Функция ИЛИ</a> | Возвращает значение `TRUE`, если по крайней мере один аргумент имеет значение ИСТИНА |
| <a href="https://support.microsoft.com/office/44f33460-5be5-4c90-b857-22308892adaf" target="_blank">Функция ПДЛИТ</a> | Возвращает количество периодов, необходимых инвестициям для достижения определенной стоимости |
| <a href="https://support.microsoft.com/office/bbaa7204-e9e1-4010-85bf-c31dc5dce4ba" target="_blank">Функция ПРОЦЕНТИЛЬ.ИСКЛ</a> | Возвращает k-ю процентиль для значений диапазона, где k — число от 0 до 1 (исключительно) |
| <a href="https://support.microsoft.com/office/680f9539-45eb-410b-9a5e-c1355e5fe2ed" target="_blank">Функция ПРОЦЕНТИЛЬ.ВКЛ</a> | Возвращает k-ю процентиль для значений диапазона |
| <a href="https://support.microsoft.com/office/d8afee96-b7e2-4a2f-8c01-8fcdedaa6314" target="_blank">Функция ПРОЦЕНТРАНГ.ИСКЛ</a> | Возвращает процентный ранг значения в наборе данных (от 0 до 1 исключительно) |
| <a href="https://support.microsoft.com/office/149592c9-00c0-49ba-86c1-c1f45b80463a" target="_blank">Функция ПРОЦЕНТРАНГ.ВКЛ</a> | Возвращает процентный ранг значения в наборе данных |
| <a href="https://support.microsoft.com/office/3bd1cb9a-2880-41ab-a197-f246a7a602d3" target="_blank">Функция ПЕРЕСТ</a> | Возвращает число перестановок для заданного количества объектов |
| <a href="https://support.microsoft.com/office/6c7d7fdc-d657-44e6-aa19-2857b25cae4e" target="_blank">Функция ПЕРЕСТА</a> | Возвращает число перестановок для заданного количества объектов (с повторами), которые можно выбрать из общего количества объектов |
| <a href="https://support.microsoft.com/office/23e49bc6-a8e8-402d-98d3-9ded87f6295c" target="_blank">Функция ФИ</a> | Возвращает значение функции плотности для стандартного нормального распределения |
| <a href="https://support.microsoft.com/office/264199d0-a3ba-46b8-975a-c4a04608989b" target="_blank">Функция ПИ</a> | Возвращает значение числа "пи" |
| <a href="https://support.microsoft.com/office/0214da64-9a63-4996-bc20-214433fa6441" target="_blank">Функция ПЛТ</a> | Возвращает сумму периодического платежа для аннуитета |
| <a href="https://support.microsoft.com/office/8fe148ff-39a2-46cb-abf3-7772695d9636" target="_blank">Функция ПУАССОН.РАСП</a> | Возвращает распределение Пуассона |
| <a href="https://support.microsoft.com/office/d3f2908b-56f4-4c3f-895a-07fb519c362a" target="_blank">Функция СТЕПЕНЬ</a> | Возвращает число, возведенное в степень |
| <a href="https://support.microsoft.com/office/c370d9e3-7749-4ca4-beea-b06c6ac95e1b" target="_blank">Функция ОСПЛТ</a> | Возвращает размер платежа для погашения основной суммы инвестиции за определенный период |
| <a href="https://support.microsoft.com/office/3ea9deac-8dfa-436f-a7c8-17ea02c21b0a" target="_blank">Функция ЦЕНА</a> | Возвращает стоимость ценной бумаги номиналом 100 рублей с периодической выплатой процентов |
| <a href="https://support.microsoft.com/office/d06ad7c1-380e-4be7-9fd9-75e3079acfd3" target="_blank">Функция ЦЕНАСКИДКА</a> | Возвращает стоимость дисконтной ценной бумаги номиналом 100 рублей |
| <a href="https://support.microsoft.com/office/52c3b4da-bc7e-476a-989f-a95f675cae77" target="_blank">Функция ЦЕНАПОГАШ</a> | Возвращает стоимость ценной бумаги номиналом 100 рублей с выплатой процентов в срок погашения |
| <a href="https://support.microsoft.com/office/8e6b5b24-90ee-4650-aeec-80982a0512ce" target="_blank">Функция ПРОИЗВЕД</a> | Возвращает произведение аргументов |
| <a href="https://support.microsoft.com/office/52a5a283-e8b2-49be-8506-b2887b889f94" target="_blank">Функция ПРОПНАЧ</a> | Преобразует первые буквы всех слов в заглавные |
| <a href="https://support.microsoft.com/office/23879d31-0e02-4321-be01-da16e8168cbd" target="_blank">Функция ПС</a> | Возвращает текущую стоимость инвестиций |
| <a href="https://support.microsoft.com/office/5a355b7a-840b-4a01-b0f1-f538c2864cad" target="_blank">Функция КВАРТИЛЬ.ИСКЛ</a> | Возвращает квартиль набора данных на основании значений процентили от 0 до 1 (исключительно) |
| <a href="https://support.microsoft.com/office/1bbacc80-5075-42f1-aed6-47d735c4819d" target="_blank">Функция КВАРТИЛЬ.ВКЛ</a> | Возвращает квартиль набора данных |
| <a href="https://support.microsoft.com/office/9f7bf099-2a18-4282-8fa4-65290cc99dee" target="_blank">Функция ЧАСТНОЕ</a> | Возвращает целую часть от деления |
| <a href="https://support.microsoft.com/office/ac409508-3d48-45f5-ac02-1497c92de5bf" target="_blank">Функция РАДИАНЫ</a> | Преобразует градусы в радианы |
| <a href="https://support.microsoft.com/office/4cbfa695-8869-4788-8d90-021ea9f5be73" target="_blank">Функция СЛЧИС</a> | Возвращает случайное число от 0 до 1 |
| <a href="https://support.microsoft.com/office/4cc7f0d1-87dc-4eb7-987f-a469ab381685" target="_blank">Функция СЛУЧМЕЖДУ</a> | Возвращает случайное число между двумя заданными числами |
| <a href="https://support.microsoft.com/office/bd406a6f-eb38-4d73-aa8e-6d1c3c72e83a" target="_blank">Функция РАНГ.СР</a> | Возвращает ранг числа в списке чисел |
| <a href="https://support.microsoft.com/office/284858ce-8ef6-450e-b662-26245be04a40" target="_blank">Функция РАНГ.РВ</a> | Возвращает ранг числа в списке чисел |
| <a href="https://support.microsoft.com/office/9f665657-4a7e-4bb7-a030-83fc59e748ce" target="_blank">Функция СТАВКА</a> | Возвращает процентную ставку по аннуитету за один период |
| <a href="https://support.microsoft.com/office/7a3f8b93-6611-4f81-8576-828312c9b5e5" target="_blank">Функция ПОЛУЧЕНО</a> | Возвращает сумму, полученную в конце срока погашения по полностью инвестированной ценной бумаге |
| <a href="https://support.microsoft.com/office/8d799074-2425-4a8a-84bc-82472868878a" target="_blank">Функции ЗАМЕНИТЬ, ЗАМЕНИТЬБ</a> | Заменяют знаки в тексте |
| <a href="https://support.microsoft.com/office/04c4d778-e712-43b4-9c15-d656582bb061" target="_blank">Функция ПОВТОР</a> | Повторяет текст заданное число раз |
| <a href="https://support.microsoft.com/office/240267ee-9afa-4639-a02b-f19e1786cf2f" target="_blank">Функции ПРАВСИМВ, ПРАВБ</a> | Возвращают последние символы в текстовой строке |
| <a href="https://support.microsoft.com/office/d6b0b99e-de46-4704-a518-b45a0f8b56f5" target="_blank">Функция РИМСКОЕ</a> | Преобразует арабское число в римское в текстовом формате |
| <a href="https://support.microsoft.com/office/c018c5d8-40fb-4053-90b1-b3e7f61a213c" target="_blank">Функция ОКРУГЛ</a> | Округляет число до указанного количества цифр |
| <a href="https://support.microsoft.com/office/2ec94c73-241f-4b01-8c6f-17e6d7968f53" target="_blank">Функция ОКРУГЛВНИЗ</a> | Округляет число к меньшему до ближайшего по модулю |
| <a href="https://support.microsoft.com/office/f8bc9b23-e795-47db-8703-db171d0c42a7" target="_blank">Функция ОКРУГЛВВЕРХ</a> | Округляет число к большему по модулю |
| <a href="https://support.microsoft.com/office/b592593e-3fc2-47f2-bec1-bda493811597" target="_blank">Функция ЧСТРОК</a> | Возвращает количество строк в ссылке |
| <a href="https://support.microsoft.com/office/6f5822d8-7ef1-4233-944c-79e8172930f4" target="_blank">Функция ЭКВ.СТАВКА</a> | Возвращает эквивалентную процентную ставку для роста инвестиций |
| <a href="https://support.microsoft.com/office/ff224717-9c87-4170-9b58-d069ced6d5f7" target="_blank">Функция SEC</a> | Возвращает секанс угла |
| <a href="https://support.microsoft.com/office/e05a789f-5ff7-4d7f-984a-5edb9b09556f" target="_blank">Функция SECH</a> | Возвращает гиперболический секанс угла |
| <a href="https://support.microsoft.com/office/740d1cfc-553c-4099-b668-80eaa24e8af1" target="_blank">Функция СЕКУНДЫ</a> | Преобразует порядковый номер в секунды |
| <a href="https://support.microsoft.com/office/a3ab25b5-1093-4f5b-b084-96c49087f637" target="_blank">Функция РЯД.СУММ</a> | Возвращает сумму степенного ряда, вычисленную по формуле |
| <a href="https://support.microsoft.com/office/44718b6f-8b87-47a1-a9d6-b701c06cff24" target="_blank">Функция ЛИСТ</a> | Возвращает номер указанного листа |
| <a href="https://support.microsoft.com/office/770515eb-e1e8-45ce-8066-b557e5e4b80b" target="_blank">Функция ЛИСТЫ</a> | Возвращает количество листов в ссылке |
| <a href="https://support.microsoft.com/office/109c932d-fcdc-4023-91f1-2dd0e916a1d8" target="_blank">Функция ЗНАК</a> | Возвращает знак числа |
| <a href="https://support.microsoft.com/office/cf0e3432-8b9e-483c-bc55-a76651c95602" target="_blank">Функция SIN</a> | Возвращает синус заданного угла |
| <a href="https://support.microsoft.com/office/1e4e8b9f-2b65-43fc-ab8a-0a37f4081fa7" target="_blank">Функция SINH</a> | Возвращает гиперболический синус числа |
| <a href="https://support.microsoft.com/office/bdf49d86-b1ef-4804-a046-28eaea69c9fa" target="_blank">Функция СКОС</a> | Возвращает асимметрию распределения |
| <a href="https://support.microsoft.com/office/76530a5c-99b9-48a1-8392-26632d542fcb" target="_blank">Функция СКОС.Г</a> | Возвращает асимметрию распределения на основании совокупности: характеристика степени асимметрии распределения относительно среднего значения |
| <a href="https://support.microsoft.com/office/cdb666e5-c1c6-40a7-806a-e695edc2f1c8" target="_blank">Функция АПЛ</a> | Возвращает сумму амортизации актива за один период, рассчитанную линейным методом |
| <a href="https://support.microsoft.com/office/17da8222-7c82-42b2-961b-14c45384df07" target="_blank">Функция НАИМЕНЬШИЙ</a> | Возвращает k-е наименьшее значение в наборе данных |
| <a href="https://support.microsoft.com/office/654975c2-05c4-4831-9a24-2c65e4040fdf" target="_blank">Функция КОРЕНЬ</a> | Возвращает положительный квадратный корень |
| <a href="https://support.microsoft.com/office/1fb4e63f-9b51-46d6-ad68-b3e7a8b519b4" target="_blank">Функция КОРЕНЬПИ</a> | Возвращает квадратный корень произведения "пи" и числа) |
| <a href="https://support.microsoft.com/office/81d66554-2d54-40ec-ba83-6437108ee775" target="_blank">Функция НОРМАЛИЗАЦИЯ</a> | Возвращает нормализованное значение |
| <a href="https://support.microsoft.com/office/6e917c05-31a0-496f-ade7-4f4e7462f285" target="_blank">Функция СТАНДОТКЛОН.Г</a> | Вычисляет стандартное отклонение для генеральной совокупности |
| <a href="https://support.microsoft.com/office/7d69cf97-0c1f-4acf-be27-f3e83904cc23" target="_blank">Функция СТАНДОТКЛОН.В</a> | Оценивает стандартное отклонение для выборки |
| <a href="https://support.microsoft.com/office/5ff38888-7ea5-48de-9a6d-11ed73b29e9d" target="_blank">Функция СТАНДОТКЛОНА</a> | Оценивает стандартное отклонение для выборки, включая числовые, текстовые и логические значения |
| <a href="https://support.microsoft.com/office/5578d4d6-455a-4308-9991-d405afe2c28c" target="_blank">Функция СТАНДОТКЛОНПА</a> | Вычисляет стандартное отклонение по генеральной совокупности, включая числовые, текстовые и логические значения |
| <a href="https://support.microsoft.com/office/6434944e-a904-4336-a9b0-1e58df3bc332" target="_blank">Функция ПОДСТАВИТЬ</a> | Заменяет один текст на другой |
| <a href="https://support.microsoft.com/office/7b027003-f060-4ade-9040-e478765b9939" target="_blank">Функция ПРОМЕЖУТОЧНЫЕ.ИТОГИ</a> | Возвращает промежуточный итог в список или базу данных |
| <a href="https://support.microsoft.com/office/043e1c7d-7726-4e80-8f32-07b23e057f89" target="_blank">Функция СУММ</a> | Суммирует аргументы |
| <a href="https://support.microsoft.com/office/169b8c99-c05c-4483-a712-1697a653039b" target="_blank">Функция СУММЕСЛИ</a> | Суммирует ячейки, соответствующие определенному условию |
| <a href="https://support.microsoft.com/office/c9e748f5-7ea7-455d-9406-611cebce642b" target="_blank">Функция СУММЕСЛИМН</a> | Суммирует ячейки в диапазоне, соответствующие нескольким условиям |
| <a href="https://support.microsoft.com/office/e3313c02-51cc-4963-aae6-31442d9ec307" target="_blank">Функция СУММКВ</a> | Возвращает сумму квадратов аргументов |
| <a href="https://support.microsoft.com/office/069f8106-b60b-4ca2-98e0-2a0f206bdb27" target="_blank">Функция АСЧ</a> | Возвращает сумму амортизации актива за указанный период, рассчитанную методом суммы годовых цифр |
| <a href="https://support.microsoft.com/office/fb83aeec-45e7-4924-af95-53e073541228" target="_blank">Функция Т</a> | Преобразует аргументы в текст |
| <a href="https://support.microsoft.com/office/4329459f-ae91-48c2-bba8-1ead1c6c21b2" target="_blank">Функция СТЬЮДЕНТ.РАСП</a> | Возвращает процентные точки (вероятность) для t-распределения Стьюдента |
| <a href="https://support.microsoft.com/office/198e9340-e360-4230-bd21-f52f22ff5c28" target="_blank">Функция СТЬЮДЕНТ.РАСП.2Х</a> | Возвращает процентные точки (вероятность) для t-распределения Стьюдента |
| <a href="https://support.microsoft.com/office/20a30020-86f9-4b35-af1f-7ef6ae683eda" target="_blank">Функция СТЬЮДЕНТ.РАСП.ПХ</a> | Возвращает t-распределение Стьюдента |
| <a href="https://support.microsoft.com/office/2908272b-4e61-4942-9df9-a25fec9b0e2e" target="_blank">Функция СТЬЮДЕНТ.ОБР</a> | Возвращает значение t для t-распределения Стьюдента как функцию вероятности и степеней свободы |
| <a href="https://support.microsoft.com/office/ce72ea19-ec6c-4be7-bed2-b9baf2264f17" target="_blank">Функция СТЬЮДЕНТ.ОБР.2Х</a> | Возвращает обратное t-распределение Стьюдента |
| <a href="https://support.microsoft.com/office/08851a40-179f-4052-b789-d7f699447401" target="_blank">Функция TAN</a> | Возвращает тангенс числа |
| <a href="https://support.microsoft.com/office/017222f0-a0c3-4f69-9787-b3202295dc6c" target="_blank">Функция TANH</a> | Возвращает гиперболический тангенс числа |
| <a href="https://support.microsoft.com/office/2ab72d90-9b4d-4efe-9fc2-0f81f2c19c8c" target="_blank">Функция РАВНОКЧЕК</a> | Возвращает облигационно-эквивалентную доходность для Казначейского векселя |
| <a href="https://support.microsoft.com/office/eacca992-c29d-425a-9eb8-0513fe6035a2" target="_blank">Функция ЦЕНАКЧЕК</a> | Возвращает цену Казначейского векселя номиналом 100 рублей |
| <a href="https://support.microsoft.com/office/6d381232-f4b0-4cd5-8e97-45b9c03468ba" target="_blank">Функция ДОХОДКЧЕК</a> | Возвращает доходность Казначейского векселя |
| <a href="https://support.microsoft.com/office/20d5ac4d-7b94-49fd-bb38-93d29371225c" target="_blank">Функция ТЕКСТ</a> | Преобразует число в текст заданного формата |
| <a href="https://support.microsoft.com/office/9a5aff99-8f7d-4611-845e-747d0b8d5457" target="_blank">Функция ВРЕМЯ</a> | Возвращает порядковый номер определенного времени |
| <a href="https://support.microsoft.com/office/0b615c12-33d8-4431-bf3d-f3eb6d186645" target="_blank">Функция ВРЕМЗНАЧ</a> | Преобразует время из текстового формата в порядковый номер |
| <a href="https://support.microsoft.com/office/5eb3078d-a82c-4736-8930-2f51a028fdd9" target="_blank">Функция СЕГОДНЯ</a> | Возвращает порядковый номер текущей даты |
| <a href="https://support.microsoft.com/office/410388fa-c5df-49c6-b16c-9e5630b479f9" target="_blank">Функция СЖПРОБЕЛЫ</a> | Удаляет из текста пробелы |
| <a href="https://support.microsoft.com/office/d90c9878-a119-4746-88fa-63d988f511d3" target="_blank">Функция УРЕЗСРЕДНЕЕ</a> | Возвращает среднее арифметическое внутренних значений набора данных |
| <a href="https://support.microsoft.com/office/7652c6e3-8987-48d0-97cd-ef223246b3fb" target="_blank">Функция ИСТИНА</a> | Возвращает логическое значение `TRUE` |
| <a href="https://support.microsoft.com/office/8b86a64c-3127-43db-ba14-aa5ceb292721" target="_blank">Функция ОТБР</a> | Усекает число до целого |
| <a href="https://support.microsoft.com/office/45b4e688-4bc3-48b3-a105-ffa892995899" target="_blank">Функция ТИП</a> | Возвращает число, обозначающее тип данных значения |
| <a href="https://support.microsoft.com/office/ffeb64f5-f131-44c6-b332-5cd72f0659b8" target="_blank">Функция ЮНИСИМВ</a> | Возвращает символ Юникод, который соответствует указанному числовому значению |
| <a href="https://support.microsoft.com/office/adb74aaa-a2a5-4dde-aff6-966e4e81f16f" target="_blank">Функция UNICODE</a> | Возвращает числовой код, который соответствует первому символу в текстовой строке |
| <a href="https://support.microsoft.com/office/c11f29b3-d1a3-4537-8df6-04d0049963d6" target="_blank">Функция ПРОПИСН</a> | Преобразует текст в верхний регистр |
| <a href="https://support.microsoft.com/office/257d0108-07dc-437d-ae1c-bc2d3953d8c2" target="_blank">Функция ЗНАЧЕН</a> | Преобразует текстовый аргумент в число |
| <a href="https://support.microsoft.com/office/73d1285c-108c-4843-ba5d-a51f90656f3a" target="_blank">Функция ДИСП.Г</a> | Вычисляет дисперсию для генеральной совокупности |
| <a href="https://support.microsoft.com/office/913633de-136b-449d-813e-65a00b2b990b" target="_blank">Функция ДИСП.В</a> | Оценивает дисперсию для выборки |
| <a href="https://support.microsoft.com/office/3de77469-fa3a-47b4-85fd-81758a1e1d07" target="_blank">Функция ДИСПА</a> | Оценивает дисперсию для выборки, включая числовые, текстовые и логические значения |
| <a href="https://support.microsoft.com/office/59a62635-4e89-4fad-88ac-ce4dc0513b96" target="_blank">Функция ДИСПРА</a> | Вычисляет дисперсию для генеральной совокупности, включая числовые, текстовые и логические значения |
| <a href="https://support.microsoft.com/office/dde4e207-f3fa-488d-91d2-66d55e861d73" target="_blank">Функция ПУО</a> | Возвращает сумму амортизации актива за указанный или неполный период, начисляемую по методу убывающего остатка |
| <a href="https://support.microsoft.com/office/0bbc8083-26fe-4963-8ab8-93a18ad188a1" target="_blank">Функция ВПР</a> | Возвращает значение ячейки первого столбца и соответствующей строки массива |
| <a href="https://support.microsoft.com/office/60e44483-2ed1-439f-8bd0-e404c190949a" target="_blank">Функция ДЕНЬНЕД</a> | Преобразует порядковый номер в день недели |
| <a href="https://support.microsoft.com/office/e5c43a03-b4ab-426c-b411-b18c13c75340" target="_blank">Функция НОМНЕДЕЛИ</a> | Преобразует порядковый номер в число, обозначающее номер недели в году |
| <a href="https://support.microsoft.com/office/4e783c39-9325-49be-bbc9-a83ef82b45db" target="_blank">Функция ВЕЙБУЛЛ.РАСП</a> | Возвращает распределение Вейбулла |
| <a href="https://support.microsoft.com/office/f764a5b7-05fc-4494-9486-60d494efbf33" target="_blank">Функция РАБДЕНЬ</a> | Возвращает дату, отстоящую на указанное количество рабочих дней от начальной даты |
| <a href="https://support.microsoft.com/office/a378391c-9ba7-4678-8a39-39611a9bf81d" target="_blank">Функция РАБДЕНЬ.МЕЖД</a> | Возвращает дату, отстоящую на указанное количество рабочих дней от начальной даты, с использованием параметров, определяющих, сколько в неделе выходных дней и какие дни являются выходными |
| <a href="https://support.microsoft.com/office/de1242ec-6477-445b-b11b-a303ad9adc9d" target="_blank">Функция ЧИСТВНДОХ</a> | Возвращает внутреннюю норму доходности на основании ряда нерегулярных выплат |
| <a href="https://support.microsoft.com/office/1b42bbf6-370f-4532-a0eb-d67c16b664b7" target="_blank">Функция ЧИСТНЗ</a> | Возвращает чистую приведенную стоимость на основании ряда нерегулярных выплат |
| <a href="https://support.microsoft.com/office/1548d4c2-5e47-4f77-9a92-0533bba14f37" target="_blank">Функция ИСКЛИЛИ</a> | Возвращает логическое исключающее ИЛИ всех аргументов |
| <a href="https://support.microsoft.com/office/c64f017a-1354-490d-981f-578e8ec8d3b9" target="_blank">Функция ГОД</a> | Преобразует порядковый номер в год |
| <a href="https://support.microsoft.com/office/3844141e-c76d-4143-82b6-208454ddc6a8" target="_blank">Функция ДОЛЯГОДА</a> | Возвращает количество лет, в том числе неполных, между начальной и конечной датами. |
| <a href="https://support.microsoft.com/office/f5f5ca43-c4bd-434f-8bd2-ed3c9727a4fe" target="_blank">Функция ДОХОД</a> | Возвращает доходность ценной бумаги с периодическими выплатами процентов |
| <a href="https://support.microsoft.com/office/a9dbdbae-7dae-46de-b995-615faffaaed7" target="_blank">Функция ДОХОДСКИДКА</a> | Возвращает годовую доходность дисконтной ценной бумаги, например Казначейского векселя |
| <a href="https://support.microsoft.com/office/ba7d1809-0d33-4bcb-96c7-6c56ec62ef6f" target="_blank">Функция ДОХОДПОГАШ</a> | Возвращает годовую доходность ценной бумаги с выплатой процентов в срок погашения |
| <a href="https://support.microsoft.com/office/d633d5a3-2031-4614-a016-92180ad82bee" target="_blank">Функция Z.ТЕСТ</a> | Возвращает одностороннее вероятностное значение z-теста |

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Класс functions (API JavaScript для Excel)](/javascript/api/excel/excel.functions)
- [Объект Функции книги (API JavaScript для Excel)](/javascript/api/excel/excel.workbook#excel-excel-workbook-functions-member)
