---
title: Задержка выполнения во время редактирования ячейки
description: Узнайте, как отложить выполнение функции Excel.run при редактировании ячейки.
ms.date: 02/16/2022
ms.localizationpriority: medium
ms.openlocfilehash: c434fddf70c89d49712c96a42db772d67168a1fb
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958539"
---
# <a name="delay-execution-while-cell-is-being-edited"></a>Задержка выполнения во время редактирования ячейки

`Excel.run` имеет перегрузку, которая принимает объект [Excel.RunOptions](/javascript/api/excel/excel.runoptions) . Он содержит набор свойств, влияющих на поведение платформы при выполнении функции. В настоящее время поддерживается следующее свойство.

- `delayForCellEdit`: определяет, откладывает ли Excel пакетный запрос до выхода пользователя из режима правки ячейки. Если `true`пакетный запрос задерживается и выполняется, когда пользователь выходит из режима редактирования ячеек. При `false`этом пакетный запрос автоматически завершается сбоем, если пользователь находится в режиме редактирования ячейки (что приводит к ошибке для доступа к пользователю). Поведение по умолчанию без указанного `delayForCellEdit` свойства эквивалентно его значению `false`.

```js
await Excel.run({ delayForCellEdit: true }, async (context) => { ... });
```
