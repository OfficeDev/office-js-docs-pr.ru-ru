---
title: Пользовательские функции и типы данных
description: Используйте типы данных Excel с пользовательскими функциями и надстройками Office.
ms.date: 10/17/2022
ms.topic: conceptual
ms.custom: scenarios:getting-started
ms.localizationpriority: medium
ms.openlocfilehash: 6ea2287dbf83a5acc45f64c6f5071e504e66bbce
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607431"
---
# <a name="use-data-types-with-custom-functions-in-excel"></a>Использование типов данных с пользовательскими функциями в Excel

Типы данных расширяют программный интерфейс API JavaScript для Excel для поддержки новых типов данных в дополнение к исходным четырем типам значений ячеек (строка, число, логическое значение и ошибка). Типы данных включают поддержку веб-изображений, отформатированных числов, сущностей и массивов в сущностях.

За счет этих типов данных расширяются возможности пользовательских функций, поскольку такие функции принимают типы данных в качестве значений как входных, так и выходных данных. Типы данных можно создавать с помощью пользовательских функций или использовать существующие типы данных в качестве аргументов функций при вычислениях. После установки схемы JSON типа данных она поддерживается во всех вычислениях.

Дополнительные сведения об использовании типов данных с помощью надстройки Excel см. в статье [Обзор типов данных в надстройках Excel](excel-data-types-overview.md).

## <a name="how-custom-functions-handle-data-types"></a>Как пользовательские функции обрабатывают типы данных

Пользовательские функции могут распознавать типы данных и принимать их в качестве значений параметров. Пользовательская функция может создать новый тип данных для возвращаемого значения. Пользовательские функции используют ту же схему JSON для типов данных, что и программный интерфейс API JavaScript для Excel. Эта схема JSON поддерживается при вычислениях и оценке пользовательскими функциями.

> [!NOTE]
> Пользовательские функции не поддерживают полную функциональность объектов расширенных ошибок, предлагаемых типами данных. Пользовательская функция может принимать объект ошибки типов данных, но он не будет поддерживаться во время вычислений. В настоящее время пользовательские функции поддерживают только ошибки, включенные в объект [CustomFunctions.Error](custom-functions-errors.md).

## <a name="enable-data-types-for-custom-functions"></a>Включение типов данных для пользовательских функций

Проекты пользовательских функций включают файл метаданных JSON. Этот файл метаданных JSON отличается от схемы JSON, используемой API-интерфейсами типов данных. Чтобы использовать интеграцию типов данных с пользовательскими функциями, файл метаданных JSON пользовательских функций должен обновляться вручную, чтобы включить свойство `allowCustomDataForDataTypeAny`. Присвойте этому свойству значение `true`.

Полное описание процесса создания метаданных JSON вручную см. в статье о создании метаданных [JSON вручную для пользовательских функций](custom-functions-json.md). Дополнительные сведения об этом свойстве см. в разделе [allowCustomDataForDataTypeAny](custom-functions-json.md#allowcustomdatafordatatypeany).

## <a name="output-a-formatted-number-value"></a>Вывод форматированного числового значения

В следующем примере кода показано, как создать тип данных [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue) с помощью пользовательской функции. Функция принимает базовое число и параметр формата в качестве входных параметров и возвращает тип данных форматированного числа в качестве выходного значения.

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

В следующем примере кода показана пользовательская функция, которая принимает тип данных [EntityCellValue](/javascript/api/excel/excel.entitycellvalue) в качестве входного значения. Если параметру `attribute` присвоено значение `text`, функция возвращает свойство `text` значения сущности. В противном случае функция возвращает свойство `basicValue` значения сущности.

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

## <a name="next-steps"></a>Дальнейшие действия

Чтобы поэкспериментировать с пользовательскими функциями и типами данных, установите [Script Lab в Excel](../overview/explore-with-script-lab.md) и попробуйте использовать типы данных [:](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/16-custom-functions/data-types-custom-functions.yaml) фрагмент пользовательских функций в нашей **библиотеке примеров**.

## <a name="see-also"></a>См. также

* [Обзор типов данных в надстройках Excel](excel-data-types-overview.md)
* [Основные понятия, связанные с типами данных Excel](excel-data-types-concepts.md)
* [Настройка надстройки Office для использования общей среды выполнения](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
