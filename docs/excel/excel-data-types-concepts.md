---
title: Основные понятия, связанные с типами данных API JavaScript для Excel
description: Информация об основных понятиях для использования типов данных Excel в надстройках Office.
ms.date: 01/14/2021
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: a769010ad46af7bba2210d9a6f9d66082cb3f815
ms.sourcegitcommit: 45f7482d5adcb779a9672669360ca4d8d5c85207
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/19/2022
ms.locfileid: "62074310"
---
# <a name="excel-data-types-core-concepts-preview"></a>Основные понятия, связанные с типами данных Excel (предварительная версия)

> [!NOTE]
> API типов данных в настоящее время можно использовать только в общедоступной предварительной версии. API предварительной версии могут быть изменены и не предназначены для использования в рабочей среде. Рекомендуется использовать их только в тестовой среде и среде разработки. Не используйте API предварительной версии в рабочей среде или в важных деловых документах.
>
> Чтобы использовать API предварительной версии:
>
> - Необходимо ссылаться на **бета-версию** библиотеки в сети доставки содержимого (CDN) (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). [Файл определения типа](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) для компиляции TypeScript и IntelliSense находится в сети CDN и имеет тип [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). Эти типы можно установить с помощью `npm install --save-dev @types/office-js-preview`. Дополнительные сведения см. в файле сведений пакета NPM [@microsoft/office-js](https://www.npmjs.com/package/@microsoft/office-js).
> - Возможно, вам потребуется присоединиться к [программе предварительной оценки Office](https://insider.office.com), чтобы получить доступ к более поздним сборкам Office.
>
> Чтобы попробовать типы данных в Office для Windows, номер вашей сборки Excel должен быть не ниже 16.0.14626.10000. Чтобы попробовать типы данных в Office для Mac, номер вашей сборки Excel должен быть не ниже 16.55.21102600.

В этой статье рассказывается о том, как использовать [API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md) для работы с типами данных. Здесь представлены основные понятия, лежащие в основе разработки типов данных.

## <a name="core-concepts"></a>Основные понятия

Используйте свойство [`Range.valuesAsJson`](/javascript/api/excel/excel.range#valuesAsJson) для работы со значениями типов данных. Это свойство аналогично свойству [Range.values](/javascript/api/excel/excel.range#values), но `Range.values` возвращает только четыре основных типа: значения строки, числа, логического типа или ошибки. Свойство `Range.valuesAsJson` может возвращать расширенную информацию об этих четырех основных типах, а также такие типы данных, как форматированное число, сущность и веб-изображение.

### <a name="json-schema"></a>Схема JSON

Каждый тип данных использует схему метаданных JSON, разработанную для этого типа. Это определяет [CellValueType](/javascript/api/excel/excel.cellvaluetype) данных и дополнительные сведения о ячейке, например `basicValue`, `numberFormat` или `address`. Каждый тип `CellValueType` имеет свойства, доступные в соответствии с этим типом. Например, тип `webImage` включает свойства [altText](/javascript/api/excel/excel.webimagecellvalue#altText) и [attribution](/javascript/api/excel/excel.webimagecellvalue#attribution). В следующих разделах приводятся примеры кода JSON для типов форматированного числа, сущности и веб-изображения.

Схема метаданных JSON для каждого типа данных также включает одно или несколько свойств только для чтения, которые используются в расчетах при обнаружении несовместимых сценариев, таких как версия Excel, которая не соответствует минимальному требованию к номеру сборки для функции типов данных. Свойство `basicType` является частью метаданных JSON каждого типа данных и всегда является свойством только для чтения. Свойство `basicType` используется в качестве резервного, если тип данных не поддерживается или имеет неправильный формат.

## <a name="formatted-number-values"></a>Форматированное число

Объект [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue) позволяет настройкам Excel определять свойство `numberFormat` для некоторого значения. После того как свойство форматированного числа присвоено значению, оно сопровождает это значение в расчетах и может возвращаться функциями.

В следующем примере кода JSON показана полная схема значения форматированного числа. Значение форматированного числа `myDate` в примере кода отображается в пользовательском интерфейсе Excel как **1/16/1990**. Если минимальные требования к совместимости для функции типов данных не выполнены, вычисления используют `basicValue` вместо форматированного числа.

```json
// This is an example of the complete JSON of a formatted number value.
// In this case, the number is formatted as a date.
const myDate = {
    type: Excel.CellValueType.formattedNumber,
    basicValue: 32889.0,
    basicType: Excel.CellValueType.double, // A readonly property. Used as a fallback in incompatible scenarios.
    numberFormat: "m/d/yyyy"
};
```

## <a name="entity-values"></a>Сущность

Значение сущности — это контейнер для типов данных, аналогичный объекту в объектно-ориентированном программировании. Сущности также поддерживают массивы в качестве свойств значения сущности. Объект [EntityCellValue](/javascript/api/excel/excel.entitycellvalue) позволяет надстройкам определять такие свойства, как `type`, `text` и `properties`. Свойство `properties` позволяет значению сущности определять и содержать дополнительные типы данных.

Свойства `basicType` и `basicValue` определяют, как вычисления читают этот тип данных сущности, если минимальные требования к совместимости для использования типов данных не выполнены. В этом сценарии этот тип данных сущности отображается как ошибка **#VALUE!** в пользовательском интерфейсе Excel.

В следующем примере кода JSON показана полная схема значения сущности, которое содержит текст, изображение, дату и дополнительное текстовое значение.

```json
// This is an example of the complete JSON for an entity value.
// The entity contains text and properties which contain an image, a date, and another text value.
const myEntity = {
    type: Excel.CellValueType.entity,
    text: "A llama",
    properties: {
        image: myImage,
        "start date": myDate,
        "quote": {
            type: Excel.CellValueType.string,
            basicValue: "I love llamas."
        }
    }, 
    basicType: Excel.CellValueType.error, // A readonly property. Used as a fallback in incompatible scenarios.
    basicValue: "#VALUE!" // A readonly property. Used as a fallback in incompatible scenarios.
};
```

## <a name="web-image-values"></a>Веб-изображение

Объект [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) создает возможность хранения изображения как части [сущности](#entity-values) или как независимого значения в диапазоне. Этот объект позволяет использовать множество свойств, включая `address`, `altText` и `relatedImagesAddress`.

Свойства `basicType` и `basicValue` определяют, как вычисления читают этот тип данных веб-изображения, если минимальные требования к совместимости для использования функции типов данных не выполнены. В этом сценарии этот тип данных веб-изображения отображается как ошибка **#VALUE!** в пользовательском интерфейсе Excel.

В следующем примере кода JSON показана полная схема веб-изображения.

```json
// This is an example of the complete JSON for a web image.
const myImage = {
    type: Excel.CellValueType.webImage,
    address: "https://bit.ly/2YGOwtw", 
    basicType: Excel.CellValueType.error, // A readonly property. Used as a fallback in incompatible scenarios.
    basicValue: "#VALUE!" // A readonly property. Used as a fallback in incompatible scenarios.
};
```

## <a name="improved-error-support"></a>Улучшенная поддержка ошибок

API типов данных представляют существующие ошибки пользовательского интерфейса Excel в качестве объектов. Теперь, когда эти ошибки доступны как объекты, надстройки могут определять или извлекать такие свойства, как `type`, `errorType` и `errorSubType`.

Ниже приводится список всех объектов ошибок с поддержкой, расширенной за счет типов данных.

- [BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue)
- [BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue)
- [CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue)
- [ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue)
- [Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue)
- [FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue)
- [GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue)
- [NotAvailableErrorCellValue](/javascript/api/excel/excel.notavailableerrorcellvalue)
- [NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)
- [NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue)
- [NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue)
- [RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue)
- [SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue)
- [ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)

Каждый из объектов ошибок может получить доступ к перечисление через свойство `errorSubType`, и в этом перечислении содержатся дополнительные данные об ошибке. Например, объект ошибки `BlockedErrorCellValue` может получить доступ к перечислению [BlockedErrorCellValueSubType](/javascript/api/excel/excel.blockederrorcellvaluesubtype). В перечислении `BlockedErrorCellValueSubType` содержатся дополнительные данные о том, что вызвало данную ошибку.

## <a name="see-also"></a>См. также

- [Обзор типов данных в надстройках Excel](excel-data-types-overview.md)
- [Справочник по API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md)
- [Пользовательские функции и типы данных](custom-functions-data-types-concepts.md)