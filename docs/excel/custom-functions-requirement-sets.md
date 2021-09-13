---
title: Наборы настраиваемой функции
description: Сведения о наборах пользовательских функций для Excel API JavaScript.
ms.date: 09/14/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 0d29d56bb41d44ed8553e97c583e41510e83c132
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151149"
---
# <a name="custom-functions-requirement-sets"></a>Наборы настраиваемой функции

[Пользовательские функции](custom-functions-overview.md) используют отдельный набор обязательных элементов из основных интерфейсов API JavaScript для Excel. В следующей таблице перечислены наборы пользовательских функций, поддерживаемые Office клиентские приложения, а также версии сборки или номер для этих приложений.

|  Набор обязательных элементов  |  Office для Windows<br>(подключено к подписке на Microsoft 365)  |  Office для iPad<br>(подключено к подписке на Microsoft 365)  |  Office для Mac<br>(подключено к подписке на Microsoft 365)  | Office в Интернете |
|:-----|-----|:-----|:-----|:-----|:-----|
| CustomFunctionsRuntime 1.3 | 16.0.13127.20296 или более поздней | Не поддерживается | 16.40.20081000 или более поздней | Июль 2020 г. |
| CustomFunctionsRuntime 1.2 | 16.0.12527.20194 или более поздней | Не поддерживается | 16.34.20020900 или более поздней | Январь 2020 г. |
| CustomFunctionsRuntime 1.1 | 16.0.12527.20092 или более поздней | Не поддерживается | 16.34 или более поздней | Май 2019 г. |

> [!NOTE]
> Excel настраиваемые функции не поддерживаются Office 2019 г. или более ранней (разовая покупка).

## <a name="customfunctionsruntime-11-12-and-13"></a>CustomFunctionsRuntime 1.1, 1.2 и 1.3

CustomFunctionsRuntime 1.1 — это первая версия API. Набор требований 1.2 добавляет объект `CustomFunctions.Error` для поддержки обработки ошибок. Набор требований 1.3 добавляет [поддержку потоковой передачи XLL](make-custom-functions-compatible-with-xll-udf.md#custom-function-behavior-for-xll-compatible-functions) и новые параметры в `ErrorCode` объект [CustomFunctions.Error.](/javascript/api/custom-functions-runtime/customfunctions.error) 

## <a name="see-also"></a>Дополнительные материалы

- [Справочная документация по настраиваемой функции](/javascript/api/custom-functions-runtime)
- [Наборы обязательных элементов API JavaScript для Excel](../reference/requirement-sets/excel-api-requirement-sets.md)
