---
title: Обзор типов данных в надстройках Excel
description: Типы данных в API JavaScript для Excel позволяют разработчикам надстроек Office работать с форматированными числами, веб-изображениями, сущностями, массивами внутри сущностей и расширенными ошибками в качестве типов данных.
ms.date: 10/14/2022
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 92f541d3b1296de5545bfb0016448f49043abcba
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607438"
---
# <a name="overview-of-data-types-in-excel-add-ins"></a>Обзор типов данных в надстройках Excel

Типы данных упорядочивают сложные структуры данных как объекты. К ним относятся отформатированные числовое значение, веб-изображения и сущности в виде [карточек сущностей](excel-data-types-entity-card.md).

До добавления типов данных в API JavaScript для Excel поддерживались строки, числа логические значения и ошибки. На уровне форматирования в пользовательском интерфейсе Excel в ячейки можно добавлять форматы валюты, даты и других видов на базе четырех исходных типов данных, но этот уровень контролирует только отображение исходных типов данных в пользовательском интерфейсе Excel. Значение числа не меняется, даже если ячейка в пользовательском интерфейсе Excel имеет формат валюты или даты. Такой разрыв между значением и форматом его отображения в пользовательском интерфейсе Excel может вести к путанице и ошибкам при вычислениях в надстройках. API типов данных являются решением этой проблемы.

Типы данных расширяют поддержку API JavaScript для Excel за пределами четырех исходных типов данных (строка, число, логическое значение [](excel-data-types-concepts.md#improved-error-support) и ошибка), включая веб-изображения[,](excel-data-types-concepts.md#web-image-values) отформатированные числовые [значения, сущности](excel-data-types-concepts.md#formatted-number-values)[,](excel-data-types-concepts.md#entity-values) массивы внутри сущностей и улучшенные типы данных ошибок в качестве гибких структур данных. Эти типы, на которых основаны различные [связанные типы данных](https://support.microsoft.com/office/what-linked-data-types-are-available-in-excel-6510ab58-52f6-4368-ba0f-6a76c0190772), делают вычисления в надстройках Excel точнее и проще и расширяют их потенциал за пределы двухмерной таблицы.

Чтобы узнать, как использовать API типов данных, начните со статьи основных понятий [типов данных Excel](excel-data-types-concepts.md) .

> [!NOTE]
> Чтобы сразу начать экспериментировать с типами данных, установите [Script Lab в Excel](../overview/explore-with-script-lab.md) и ознакомьтесь с разделом "Типы данных" в нашей **библиотеке примеров**. Вы также можете изучить Script Lab в нашем репозитории [OfficeDev/office-js-snippets](https://github.com/OfficeDev/office-js-snippets/tree/prod/samples/excel/20-data-types).

## <a name="data-types-and-custom-functions"></a>Типы данных и пользовательские функции

Типы данных делают пользовательские функции полезнее. Пользовательские функции принимают различные типы данных на вход и используют их на выходе; кроме того, в них применяется та же схема JSON для типов данных, что и в API JavaScript для Excel. На этой схеме JSON типов данных основаны все расчеты и вычисления пользовательских функций. Чтобы узнать больше об интеграции типов данных с пользовательскими функциями, см. [Пользовательские функции и типы данных](custom-functions-data-types-concepts.md).

## <a name="see-also"></a>См. также

- [Ключевые понятия типов данных в Excel](excel-data-types-concepts.md)
- [Использование карточек с типами данных значений сущностей](excel-data-types-entity-card.md)
- [Пользовательские функции и типы данных](custom-functions-data-types-concepts.md)
- [Создание и изучение типов данных в Excel](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-data-types-explorer)
- [Справочник по API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md)