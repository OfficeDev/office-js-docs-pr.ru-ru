---
title: Задержка выполнения во время редактирования ячейки
description: Узнайте, как отсрочить выполнение метода Excel.run при редактировании ячейки.
ms.date: 02/16/2022
ms.localizationpriority: medium
ms.openlocfilehash: c5609fbb2a39d6ecc69063d4bccdfbc1da1c102d
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340808"
---
# <a name="delay-execution-while-cell-is-being-edited"></a>Задержка выполнения во время редактирования ячейки

`Excel.run`имеет перегрузку, которая принимает в [Excel. Объект RunOptions](/javascript/api/excel/excel.runoptions). Он содержит набор свойств, влияющих на поведение платформы при выполнении функции. В настоящее время поддерживается следующее свойство.

- `delayForCellEdit`: определяет, откладывает ли Excel пакетный запрос до выхода пользователя из режима правки ячейки. Если присвоено значение **true**, пакетный запрос откладывается и запускается, когда пользователь выходит из режима правки ячейки. Если присвоено значение **false**, происходит автоматический сбой пакетного запроса, если пользователь находится в режиме правки ячейки (приводит к ошибке обращения к пользователю). Поведение по умолчанию при отсутствии заданного свойства `delayForCellEdit` аналогично поведению при значении **false**.

```js
await Excel.run({ delayForCellEdit: true }, async (context) => { ... });
```
