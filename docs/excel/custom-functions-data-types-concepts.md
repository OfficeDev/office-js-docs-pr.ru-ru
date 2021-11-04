---
title: Настраиваемые основные понятия функций и типов данных
description: Узнайте основные концепции использования типов Excel с помощью настраиваемой функции.
ms.date: 11/03/2021
ms.topic: conceptual
ms.custom: scenarios:getting-started
ms.localizationpriority: medium
ms.openlocfilehash: 3b7e735f78ca7b6dcdffa3bd5e8ba9c9d3093766
ms.sourcegitcommit: ad5d7ab21f64012543fb2bd9226d90330d25468b
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/04/2021
ms.locfileid: "60749408"
---
# <a name="custom-functions-and-data-types-core-concepts-preview"></a>Настраиваемые основные концепции функций и типов данных (предварительный просмотр)

[!include[Custom functions and data types availability note](../includes/excel-custom-functions-data-types-note.md)]

Типы данных повышают Excel API JavaScript, расширяя поддержку типов данных за пределами исходных четырех (строка, число, boolean и ошибка). Типы данных включают поддержку отформатированные значения номеров, веб-изображений, значений сущности и массивов в значениях сущности. Настраиваемые функции принимают типы данных в качестве значений ввода и вывода, расширяя возможности вычисления пользовательских функций.

Дополнительные сведения об использовании типов данных с помощью надстройки Excel см. в Excel [основных понятий типов данных.](excel-data-types-concepts.md)

## <a name="how-custom-functions-handle-data-types"></a>Обработка пользовательских функций типами данных

Настраиваемые функции могут распознавать типы данных и принимать их в качестве значений параметров. Настраиваемая функция может создать новый тип данных для возвращаемого значения. Пользовательские функции используют ту же схему JSON для типов данных, Excel API JavaScript, и эта схема JSON поддерживается в качестве настраиваемой функции, вычисляемой и оцениваемой.

> [!NOTE]
> Настраиваемые функции не поддерживают полную функциональность объектов расширенной ошибки, предлагаемых типами данных. Настраиваемая функция может принимать объект ошибок типов данных, но он не будет поддерживаться во время вычислений. В настоящее время пользовательские функции поддерживают только ошибки, включенные в объект [CustomFunctions.Error.](custom-functions-errors.md)

## <a name="enable-data-types-for-custom-functions"></a>Включить типы данных для настраиваемой функции

Чтобы использовать эту функцию, необходимо вручную обновить метаданные JSON. Для временного тестирования можно настроить параметры Script Lab, а не вручную обновлять метаданные JSON. В следующих разделах описаны эти действия более подробно.

### <a name="manually-update-json-metadata"></a>Ручное обновление метаданных JSON

Проекты пользовательских функций включают файл метаданных JSON. Этот файл метаданных JSON отличается от схемы JSON, используемой API типов данных. Чтобы использовать интеграцию типов данных с настраиваемой функцией, пользовательский файл метаданных JSON должен обновляться вручную, чтобы включить свойство `allowCustomDataForDataTypeAny` . Установите это свойство `true` .

Полное описание процесса создания JSON вручную см. в инструкции по созданию метаданных JSON вручную [для пользовательских функций.](custom-functions-json.md) Дополнительные сведения об этом свойстве см. в материале [allowCustomDataForDataTypeAny.](custom-functions-json.md#allowcustomdatafordatatypeany-preview)

### <a name="script-lab-option"></a>Script Lab

Настраиваемая интеграция функций с типами данных доступна для тестирования с помощью Script Lab, в дополнение к обновлению метаданных JSON, описанного в предыдущем разделе. Дополнительные дополнительные Script Lab см. в [Office API JavaScript с Script Lab.](../overview/explore-with-script-lab.md) Чтобы проверить эту функцию с помощью Script Lab, обновим параметры с помощью следующих действий.

1. Откройте области задач Script Lab **кода.**
1. В правом нижнем углу выберите кнопку **Параметры.**
1. Перейдите на **вкладку Параметры** пользователя и введите `allowCustomDataForDataTypeAny: true` .

![Снимок экрана, показывающий действия, позволяющие включить типы данных для настраиваемой функции в Script Lab.](../images/custom-functions-script-lab-data-type.png)

## <a name="output-a-formatted-number-value"></a>Вывод отформатированного значения числа

В следующем примере кода показано, как создать тип [данных FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue) с настраиваемой функцией. Функция принимает базовое число и параметр формата в качестве параметров ввода и возвращает форматированный тип данных значения номера в качестве вывода.

```js
/**
 * Take a number as the input value and return a formatted number value as the output.
 * @customfunction
 * @param {number} value
 * @param {string} format (e.g. "0.00%")
 * @returns A formatted number value.
 */
function createFormattedNumber(value, format) {
    return {
        type: "FormattedNumber",
        basicValue: value,
        numberFormat: format
    }
}
```

## <a name="input-an-entity-value"></a>Ввод значения сущности

В следующем примере кода показана настраиваемая функция, которая принимает тип данных [EntityCellValue](/javascript/api/excel/excel.entitycellvalue) в качестве ввода. Если параметр задан, функция возвращает свойство значения `attribute` `text` `text` сущности. В противном случае функция возвращает свойство значения `basicValue` сущности.

```js
/**
 * Accept an entity value data type as a function input.
 * @customfunction
 * @param {any} value
 * @param {string} attribute
 * @returns {any} The text value of the entity.
 */
function getEntityAttribute(value, attribute) {
    if (value.type == "Entity") {
        if (attribute == "text") {
            return value.text;
        } else {
            return value.properties[attribute].basicValue;
        }
    } else {
        return JSON.stringify(value);
    }
}
```

## <a name="see-also"></a>См. также

* [Обзор пользовательских функций и типов данных](custom-functions-data-types-overview.md)
* [Обзор типов данных в Excel надстройки](excel-data-types-overview.md)
* [Excel типов данных основные понятия](excel-data-types-concepts.md)
* [Настройка надстройки Office для использования общей среды выполнения JavaScript](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
