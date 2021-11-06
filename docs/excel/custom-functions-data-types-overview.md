---
title: Обзор пользовательских функций и типов данных
description: Используйте типы данных Excel с пользовательскими функциями и надстройками Office.
ms.date: 11/03/2021
ms.topic: conceptual
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 91d2fb21aae57ed7a5777136f3c4540925f339c8
ms.sourcegitcommit: 210251da940964b9eb28f1071977ea1fe80271b4
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/05/2021
ms.locfileid: "60793577"
---
# <a name="use-data-types-with-custom-functions-in-excel-preview"></a>Используйте типы данных в пользовательских функциях Excel (предварительная версия)

[!include[Custom functions and data types availability note](../includes/excel-custom-functions-data-types-note.md)]

Типы данных расширяют API JavaScript для Excel, обеспечивая поддержку типов данных, не входящих в четыре первоначальных типа данных (строка, число, логический и ошибка). Типы данных включают поддержку веб-изображений, форматированных чисел, сущностей и массивов сущностей.

За счет этих типов данных расширяются возможности пользовательских функций, поскольку такие функции принимают типы данных в качестве значений как входных, так и выходных данных. Типы данных можно создавать с помощью пользовательских функций или использовать существующие типы данных в качестве аргументов функций при вычислениях. После задания схемы JSON для типа данных эта схема сохраняется во всех вычислениях пользовательских функций.

Дополнительную информацию об использовании типов данных в надстройках Excel см. статью [Обзор типов данных в надстройках Excel](excel-data-types-overview.md). Дополнительную информацию об интеграции пользовательских типов данных в пользовательские функции см. статью [Основные понятия, связанные с пользовательскими функциями и типами данных](custom-functions-data-types-concepts.md).

## <a name="see-also"></a>См. также

* [Обзор типов данных в надстройках Excel](excel-data-types-overview.md)
* [Основные понятия, связанные с типами данных Excel](excel-data-types-concepts.md)
* [Основные понятия, связанные с пользовательскими функциями и типами данных](custom-functions-data-types-concepts.md)
* [Настройка надстройки Office для использования общей среды выполнения JavaScript](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
