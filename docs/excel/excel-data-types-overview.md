---
title: Обзор типов данных в надстройках Excel
description: Типы данных в API JavaScript для Excel позволяют разработчикам надстроек Office работать с отформатированными значениями чисел, веб-изображениями, значениями сущностей, массивами в значениях сущностей и расширенными ошибками в качестве типов.
ms.date: 11/03/2021
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 5ff0d5a055c74eeff096d45ddb6c417615775431
ms.sourcegitcommit: ad5d7ab21f64012543fb2bd9226d90330d25468b
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/04/2021
ms.locfileid: "60749394"
---
# <a name="overview-of-data-types-in-excel-add-ins-preview"></a>Обзор типов данных в надстройках Excel (предварительная версия)

> [!NOTE]
> API типов данных в настоящее время можно использовать только в общедоступной предварительной версии. API предварительной версии могут быть изменены и не предназначены для использования в рабочей среде. Не используйте API предварительной версии в рабочей среде или в важных деловых документах.

> [!IMPORTANT]
> Некоторые API типов данных, например `Range.valuesAsJSON`, находятся в активной разработке, и их общедоступная предварительная версия пока отсутствует. Эта статья служит введением в понятийный аппарат. Описанные в этой статье понятия, которые еще не используются в общедоступной предварительной версии, скоро будут опубликованы в составе этой версии.

Типы данных в API JavaScript для Excel позволяют разработчикам надстроек организовывать сложные структуры данных в качестве объектов, таких как отформатированные значения чисел, веб-изображения и значения сущностей.

До добавления типов данных в API JavaScript для Excel поддерживались строки, числа логические значения и ошибки. На уровне форматирования в пользовательском интерфейсе Excel в ячейки можно добавлять форматы валюты, даты и других видов на базе четырех исходных типов данных, но этот уровень контролирует только отображение исходных типов данных в пользовательском интерфейсе Excel. Значение числа не меняется, даже если ячейка в пользовательском интерфейсе Excel имеет формат валюты или даты. Такой разрыв между значением и форматом его отображения в пользовательском интерфейсе Excel может вести к путанице и ошибкам при вычислениях в надстройках. Решением этой проблемы являются настраиваемые типы данных.

Концепция типов данных дополняет исходные четыре типа API JavaScript в Excel (строка, число, логическое значение и ошибка) такими вариантами, как веб-изображения, отформатированные значения чисел, значения сущностей, массивы в значениях сущностей и улучшенные типы данных ошибок в виде гибких структур. Эти типы, на которых основаны различные [связанные типы данных](https://support.microsoft.com/office/what-linked-data-types-are-available-in-excel-6510ab58-52f6-4368-ba0f-6a76c0190772), делают вычисления в надстройках Excel точнее и проще и расширяют их потенциал за пределы двухмерной таблицы.

## <a name="data-types-and-custom-functions"></a>Типы данных и пользовательские функции

[!include[Custom functions and data types availability note](../includes/excel-custom-functions-data-types-note.md)]

Типы данных делают пользовательские функции полезнее. Пользовательские функции принимают различные типы данных на вход и используют их на выходе; кроме того, в них применяется та же схема JSON для типов данных, что и в API JavaScript для Excel. На этой схеме JSON типов данных основаны все расчеты и вычисления пользовательских функций. Дополнительные сведения об интеграции типов данных с пользовательскими функциями см. в статье о [ключевых понятиях пользовательских функций и типов данных.](custom-functions-data-types-concepts.md)

## <a name="see-also"></a>См. также

* [Ключевые понятия типов данных в Excel](excel-data-types-concepts.md)
* [Справочник по API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md)
* [Обзор пользовательских функций и типов данных](custom-functions-data-types-overview.md)