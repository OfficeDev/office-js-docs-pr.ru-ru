---
title: Углубленные принципы программирования с использованием интерфейса API JavaScript для Excel
description: ''
ms.date: 10/03/2018
ms.openlocfilehash: 190eb65e45ce246009b6d85d378571bd2f451e0b
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459254"
---
# <a name="advanced-programming-concepts-with-the-excel-javascript-api"></a>Углубленные принципы программирования с использованием интерфейса API JavaScript для Excel

Эта статья основана на информации, содержащейся в статье [Основные принципы программирования с использованием интерфейса API JavaScript для Excel](excel-add-ins-core-concepts.md) , и описывает более углубленные принципы создания сложных надстроек для Excel 2016 или более поздних версий.

## <a name="officejs-apis-for-excel"></a>Интерфейсы API Office.js для Excel

Надстройка Excel взаимодействует с объектами в Excel с помощью API JavaScript для Office, включающего две объектных модели JavaScript:

* **API JavaScript для Excel**. Впервые представленный в Office 2016 [, интерфейс API JavaScript для Excel](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview?view=office-js) предоставляет строго типизированные объекты, с помощью которых можно получать доступ к листам, диапазонам, таблицам, диаграммам и другим объектам. 

* **Общие API**. Впервые представленные в Office 2013, [общие API](https://docs.microsoft.com/javascript/office/javascript-api-for-office?view=office-js) можно использовать для доступа к таким компонентам, как пользовательский интерфейс, диалоговые окна и параметры клиентов, общие для нескольких типов ведущих приложений, например, Word, Excel и PowerPoint.

Хотя, скорее всего, вы будете использовать API JavaScript для Excel для разработки большинства функций, предназначенных для надстроек Excel 2016 или более поздней версии, но объекты в общих API Shared тоже будут нужны. Например:

- [Контекст](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js): объект **Context** представляет собой среду выполнения надстройки и предоставляет доступ к ключевым объектам API-интерфейса. Он содержит сведения о конфигурации книги, например, `contentLanguage` и `officeTheme` , а также сведения о среде выполнения надстройки, например, `host` и `platform`. Кроме того, объект предусматривает метод  `requirements.isSetSupported()` , который можно использовать для проверки поддержки указанного набора обязательных элементов приложением Excel, в котором установлена надстройка. 

- [Документ](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js). Объект **Document** предоставляет метод `getFileAsync()`, позволяющий скачать файл Excel, в котором работает надстройка. 

## <a name="requirement-sets"></a>Наборы обязательных элементов

Наборы обязательных элементов — это именованные группы элементов API. Надстройка Office может выполнять проверку в среде выполнения или использовать указанные в манифесте наборы обязательных элементов, чтобы определить, поддерживает ли ведущее приложение Office необходимые надстройке API. Информацию о том, как задать конкретные наборы обязательных элементов, доступные на каждой поддерживаемые платформы, см. в статье [Наборы обязательных элементов API JavaScript для Excel](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js).

### <a name="checking-for-requirement-set-support-at-runtime"></a>Проверка поддержки наборов обязательных элементов в среде выполнения

В приведенном ниже примере кода показано, как определить, поддерживает ли ведущее приложение надстройки указанный набор обязательных элементов API.

```js
if (Office.context.requirements.isSetSupported('ExcelApi', 1.3) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a>Определение поддержки наборов обязательных элементов в манифесте

|||UNTRANSLATED_CONTENT_START|||You can use the [Requirements element](https://docs.microsoft.com/javascript/office/manifest/requirements?view=office-js) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate. If the Office host or platform doesn't support the requirement sets or API methods that are specified in the **Requirements** element of the manifest, the add-in won't run in that host or platform, and won't display in the list of add-ins that are shown in **My Add-ins**.|||UNTRANSLATED_CONTENT_END||| 

В приведенном ниже примере кода показан элемент **Requirements**в манифесте надстройки, где указано, что надстройка должна загружаться во всех ведущих приложениях Office, поддерживающих набор обязательных элементов ExcelApi версии 1.3 или выше.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> Чтобы надстройка была доступна на всех платформах ведущего приложения Office, например, Excel для Windows, Excel Online и Excel для iPad, рекомендуем проверять поддержку обязательных элементов в среде выполнения, а не определять поддержку набора обязательных элементов в манифесте.

### <a name="requirement-sets-for-the-officejs-common-api"></a>Наборы обязательных элементов общего API JavaScript для Office

Сведения об общих наборах обязательных элементов API см. в статье [Общие наборы обязательных элементов API для Office](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets?view=office-js).

## <a name="loading-the-properties-of-an-object"></a>Загрузка свойств объекта

При вызове метода `load()` в объекте JavaScript для Excel интерфейс API загружает объект в память JavaScript при выполнении метода`sync()` .  Метод`load()` принимает строку, содержащую разделенные запятыми имена свойств, которые требуется загрузить, или объект, указывающий загружаемые свойства, параметры разбивки на страницы и т. д. 

> [!NOTE]
> Если вызвать метод `load()` для объекта (или коллекции), не указывая параметры, то будут загружены все скалярные свойства объекта (или все скалярные свойства всех объектов в коллекции). Чтобы уменьшить объем передаваемых данных между ведущим приложением Excel и надстройкой, не рекомендуется вызывать метод `load()` без прямого указания свойств, которые необходимо загрузить.

### <a name="method-details"></a>Сведения о методе

#### <a name="loadparam-object"></a>load(param: объект)

Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметрах.

#### <a name="syntax"></a>Синтаксис

```js
object.load(param);
```

#### <a name="parameters"></a>Параметры

|**Параметр**|**Тип**|**Описание**|
|:------------|:-------|:----------|
|`param`|объект|Необязательный атрибут. Принимает имена параметров и связей в виде строки с разделителями-запятыми или массива. Кроме того, можно передать объект, чтобы задать свойства выделения и навигации (как показано в приведенном ниже примере).|

#### <a name="returns"></a>Возвращаемое значение

пустое

#### <a name="example"></a>Пример

В приведенном ниже примере кода свойства одного диапазона Excel заданы путем копирования свойств другого диапазона. Обратите внимание, что исходный объект необходимо сначала загрузить, и только после этого можно получить доступ к значениям его свойств и записать их в указанный диапазон. В этом примере предполагается, что два диапазона (**B2:E2** и **B7:E7**) содержат данные, а их форматирование изначально отличается.

```js
Excel.run(function (ctx) { 
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var sourceRange = sheet.getRange("B2:E2");
    sourceRange.load("format/fill/color, format/font/name, format/font/color");

    return ctx.sync()
        .then(function () {
            var targetRange = sheet.getRange("B7:E7");
            targetRange.set(sourceRange); 
            targetRange.format.autofitColumns();

            return ctx.sync();
        });
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="load-option-properties"></a>Загрузка свойств параметров

Вместо того чтобы передавать строку с разделителями-запятыми или массив при вызове метода `load()`, можно передать объект, содержащий указанные ниже свойства. 

|**Свойство**|**Тип**|**Описание**|
|:-----------|:-------|:----------|
|`select`|объект|Содержит массив или разделенный запятыми список имен параметров и связей. Необязательный параметр.|
|`expand`|объект|Содержит массив или разделенный запятыми список имен связей. Необязательный параметр.|
|`top`|int| Указывает максимальное число элементов в коллекции, которые можно включить в результат. Необязательный параметр. Его можно применять, только если используется параметр нотации объектов.|
|`skip`|int|Укажите количество элементов в коллекции, которые необходимо пропустить и исключить из результата. Если указан параметр `top`, набор результатов начнется после пропуска заданного числа элементов. Необязательный параметр. Его можно применять, только если используется параметр нотации объектов.|

В приведенном ниже примере кода показано, как загрузить коллекцию листов, выбрав свойства `name` и `address` используемого диапазона для каждого листа в коллекции. В нем также указано, что следует загружать только пять первых листов в коллекции. Для обработки следующих пяти листов можно указать `top: 10` и `skip: 5` как значения атрибута. 

```js 
myWorksheets.load({
    select: 'name, userRange/address',
    expand: 'tables',
    top: 5,
    skip: 0
});
```

## <a name="scalar-and-navigation-properties"></a>Скалярные и навигационные свойства 

В документации "Справочник по API JavaScript для Excel" можно заметить, что объекты сгруппированы в две категории: **свойства** и **связи**. Свойство объекта — это скалярный элемент, например, строка, целое число или логическое значение, а связь объекта (другое название — свойство навигации) — это элемент, представляющий собой объект или их коллекцию. Например,  `name` и `position` в объекте [Worksheet](https://docs.microsoft.com/javascript/api/excel/excel.worksheet?view=office-js) являются скалярными свойствами, а `protection` и `tables` — связями (навигационные свойства). 

### <a name="scalar-properties-and-navigation-properties-with-objectload"></a>Скалярные и навигационные свойства с методом `object.load()`

При вызове метода`object.load()` без указания параметров загружаются все скалярные свойства объекта; навигационные свойства объекта не будут загружены. Кроме того, навигационные свойства невозможно загрузить напрямую. Вместо этого следует использовать метод `load()`, чтобы ссылаться на отдельные скалярные свойства в нужном свойстве навигации. Например, чтобы загрузить имя шрифта для диапазона, необходимо указать свойства навигации **format** и **font**в качестве пути к свойству **name**:

```js
someRange.load("format/font/name")
```

> [!NOTE]
> С помощью API JavaScript для Excel можно задавать скалярные свойства из навигационного свойства по пути к ним. Например, можно задать размер шрифта для диапазона с помощью`someRange.format.font.size = 10;`. Чтобы задать свойство, необязательно его загружать. 

## <a name="setting-properties-of-an-object"></a>Установка свойств объекта

Установка свойств объекта с вложенными навигационными свойствами может быть трудоемкой задачей. Вместо того чтобы задавать отдельные свойства с помощью путей навигации, как описано выше, вы можете использовать метод `object.set()`, доступный для всех объектов в API JavaScript для Excel. С помощью этого метода можно задать сразу несколько свойств объекта, передавая другой объект того же типа Office.js или объект JavaScript со свойствами, сходными по структуре со свойствами объекта, для которого вызывается метод.

> [!NOTE]
> Метод `set()` реализован только для объектов API JavaScript для Office в определенных ведущих приложениях, таких как API JavaScript для Excel. Общие API не поддерживают этот метод. 

### <a name="set-properties-object-options-object"></a>set (properties: объект, options: объект)

Свойствам объекта, для которого вызывается метод, присваиваются те же значения, что и соответствующим свойствам переданного объекта. Если параметр `properties` является объектом JavaScript, любое свойство в переданном объекте, соответствующее нередактируемому свойству в объекте, для которого вызывается метод, либо игнорируется, либо приводит к возникновению исключения, в зависимости от значения параметра `options`.

#### <a name="syntax"></a>Синтаксис

```js
object.set(properties[, options]);
```

#### <a name="parameters"></a>Параметры

|**Параметр**|**Тип**|**Описание**|
|:------------|:--------|:----------|
|`properties`|объект|Либо объект того же типа Office.js, что и объект, для которого вызывается метод, либо объект JavaScript, имена и типы свойств которого повторяют структуру объекта, для которого вызывается метод.|
|`options`|объект|Необязательный параметр. Может передаваться, только если первый параметр является объектом JavaScript. Объект может содержать следующее свойство: `throwOnReadOnly?: boolean` (по умолчанию — `true`: если переданный объект JavaScript включает нередактируемые свойства, возникает ошибка.)|

#### <a name="returns"></a>Возвращаемое значение

пустое    

#### <a name="example"></a>Пример

В приведенном ниже примере кода показано, как задать несколько свойств формата диапазона, вызвав метод`set()` и передав в него объект JavaScript, имена и типы свойств которого повторяют структуру свойств объекта **Range**. В этом примере предполагается, что данные находятся в диапазоне**B2:E2**.

```js
Excel.run(function (ctx) { 
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E2");
    range.set({
        format: {
            fill: {
                color: '#4472C4'
            },
            font: {
                name: 'Verdana',
                color: 'white'
            }
        }
    });
    range.format.autofitColumns();

    return ctx.sync(); 
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
## <a name="42ornullobject-methods"></a>Методы *OrNullObject

Многие методы API JavaScript для Excel возвращают исключение, если условие API не соблюдается. Например, если для получения листа указать имя листа, не существующее в книге, то метод`getItem()` вернет исключение`ItemNotFound`. 

Вместо того чтобы реализовывать сложную логику обработки исключений для такого сценария, можно использовать вариант метода `*OrNullObject`, доступный для нескольких методов в API JavaScript для Excel. Если указанный элемент не существует, метод `*OrNullObject` возвращает нулевой объект (не объект JavaScript`null`), вместо того чтобы возвращать исключение. Например, вы можете вызвать метод `getItemOrNullObject()` для коллекции, например, **Worksheets**, чтобы попробовать получить элемент из коллекции. Метод `getItemOrNullObject()` возвращает указанный элемент, если он существует. В противном случае возвращается нулевой объект. Возвращаемый нулевой объект содержит логическое свойство`isNullObject`, с помощью которого можно определить, существует ли объект.

В приведенном ниже примере кода осуществляется попытка получить лист "Data" (Данные) с помощью метода `getItemOrNullObject()`. Если метод возвращает нулевой объект, то, прежде чем выполнять какие-либо действия с листом, его необходимо создать.

```js
var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data"); 

return context.sync()
  .then(function() {
    if (dataSheet.isNullObject) { 
        // Create the sheet
    }

    dataSheet.position = 1;
    //...
  })
```

## <a name="see-also"></a>См. также
 
* [Основные принципы программирования с использованием интерфейса API JavaScript для Excel](excel-add-ins-core-concepts.md)
* [Примеры кода надстроек Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Оптимизация производительности API JavaScript для Excel](performance.md)
* [Справочник по API JavaScript для Excel](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview?view=office-js)
