---
title: Наборы требований к настраиваемым функциям
description: Сведения о требованиях к настраиваемым функциям для API JavaScript для Excel.
ms.date: 09/14/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 0860dd2d1b55376a85eadf04898d288d83b0205d
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819527"
---
# <a name="custom-functions-requirement-sets"></a>Наборы требований к настраиваемым функциям

[Пользовательские функции](custom-functions-overview.md) используют отдельный набор обязательных элементов из основных интерфейсов API JavaScript для Excel. В приведенной ниже таблице перечислены наборы требований пользовательских функций, Поддерживаемые клиентские приложения Office и версии сборок или номера для этих приложений.

|  Набор обязательных элементов  |  Office для Windows<br>(подключено к подписке на Microsoft 365)  |  Office для iPad<br>(подключено к подписке на Microsoft 365)  |  Office для Mac<br>(подключено к подписке на Microsoft 365)  | Office в Интернете |
|:-----|-----|:-----|:-----|:-----|:-----|
| Кустомфунктионсрунтиме 1,3 | 16.0.13127.20296 или более поздняя версия | Не поддерживается | 16.40.20081000 или более поздняя версия | Июль 2020 г. |
| Кустомфунктионсрунтиме 1,2 | 16.0.12527.20194 или более поздняя версия | Не поддерживается | 16.34.20020900 или более поздняя версия | Январь 2020 г. |
| CustomFunctionsRuntime 1.1 | 16.0.12527.20092 или более поздняя версия | Не поддерживается | 16,34 или более поздняя версия | Май 2019 г. |

> [!NOTE]
> Пользовательские функции Excel не поддерживаются в Office 2019 или более ранней версии (одноразовая покупка).

## <a name="customfunctionsruntime-11-12-and-13"></a>Кустомфунктионсрунтиме 1,1, 1,2 и 1,3

Кустомфунктионсрунтиме 1,1 — это первая версия API. Набор требований 1,2 добавляет `CustomFunctions.Error` объект для поддержки обработки ошибок. Набор обязательных элементов 1,3 добавляет поддержку [потоковой передачи XLL](make-custom-functions-compatible-with-xll-udf.md#custom-function-behavior-for-xll-compatible-functions) и новые `ErrorCode` Параметры в объект [CustomFunctions. Error](/javascript/api/custom-functions-runtime/customfunctions.error) . 

## <a name="see-also"></a>См. также

- [Справочная документация по настраиваемым функциям](/javascript/api/custom-functions-runtime)
- [Наборы обязательных элементов API JavaScript для Excel](../reference/requirement-sets/excel-api-requirement-sets.md)
