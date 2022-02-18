---
title: Наборы настраиваемой функции
description: Сведения о наборах пользовательских функций для Excel API JavaScript.
ms.date: 02/15/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 7558035b6b151977e985ec04ed1fa84c116f0886
ms.sourcegitcommit: 789545a81bd61ec2e7adef2bc24c06b5be113b00
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/18/2022
ms.locfileid: "62892519"
---
# <a name="custom-functions-requirement-sets"></a>Наборы настраиваемой функции

[Пользовательские функции](../../excel/custom-functions-overview.md) используют отдельный набор обязательных элементов из основных интерфейсов API JavaScript для Excel. В следующей таблице перечислены наборы пользовательских функций, поддерживаемые Office клиентские приложения, а также версии сборки или номер для этих приложений.

|  Набор обязательных элементов  |  Office 2021 или более поздней Windows<br>(единовременная покупка)  |  Office для Windows<br>(подключено к подписке на Microsoft 365)  |  Office для iPad<br>(подключено к подписке на Microsoft 365)  |  Office для Mac<br>(обе подписки<br> и разовая покупка Office Mac 2021 и более поздних периодов)  | Office в Интернете |
|:-----|:-----|:-----|:-----|:-----|:-----|
| CustomFunctionsRuntime 1.3 | 16.0.14326.20454 или более поздней | 16.0.13127.20296 или более поздней | Не поддерживается | 16.40.20081000 или более поздней | Июль 2020 г. |
| CustomFunctionsRuntime 1.2 | 16.0.14326.20454 или более поздней | 16.0.12527.20194 или более поздней | Не поддерживается | 16.34.20020900 или более поздней | Январь 2020 г. |
| CustomFunctionsRuntime 1.1 | 16.0.14326.20454 или более поздней | 16.0.12527.20092 или более поздней | Не поддерживается | 16.34 или более поздней | Май 2019 г. |

## <a name="customfunctionsruntime-11-12-and-13"></a>CustomFunctionsRuntime 1.1, 1.2 и 1.3

CustomFunctionsRuntime 1.1 — это первая версия API. Набор требований 1.2 добавляет объект для `CustomFunctions.Error` поддержки обработки ошибок. Набор требований 1.3 добавляет [поддержку потоковой передачи XLL](../../excel/make-custom-functions-compatible-with-xll-udf.md#custom-function-behavior-for-xll-compatible-functions) `ErrorCode` и новые параметры в [объект CustomFunctions.Error](/javascript/api/custom-functions-runtime/customfunctions.error) .

## <a name="see-also"></a>См. также

- [Справочная документация по настраиваемой функции](/javascript/api/custom-functions-runtime)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
