---
title: Задержка выполнения во время редактирования ячейки
description: Узнайте, как отсрочить выполнение метода Excel.run при редактировании ячейки.
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: 41bbfba3894bcef0c1fd075ce76557dfdc4ba4721b7bc7b19ca21756b86ccc4d
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57084287"
---
# <a name="delay-execution-while-cell-is-being-edited"></a>Задержка выполнения во время редактирования ячейки

`Excel.run`имеет перегрузку, которая принимает в [Excel. Объект RunOptions.](/javascript/api/excel/excel.runoptions) Он содержит набор свойств, влияющих на поведение платформы при выполнении функции. В настоящее время поддерживается следующее свойство.

- `delayForCellEdit`: определяет, откладывает ли Excel пакетный запрос до выхода пользователя из режима правки ячейки. Если присвоено значение **true**, пакетный запрос откладывается и запускается, когда пользователь выходит из режима правки ячейки. Если присвоено значение **false**, происходит автоматический сбой пакетного запроса, если пользователь находится в режиме правки ячейки (приводит к ошибке обращения к пользователю). Поведение по умолчанию при отсутствии заданного свойства `delayForCellEdit` аналогично поведению при значении **false**.

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```
