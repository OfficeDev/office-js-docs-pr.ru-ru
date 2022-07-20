---
title: Задержка выполнения во время редактирования ячейки
description: Узнайте, как отложить выполнение метода Excel.run при редактировании ячейки.
ms.date: 02/16/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1abcdb382150db486033b32d2521207ab0b7f28f
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889221"
---
# <a name="delay-execution-while-cell-is-being-edited"></a>Задержка выполнения во время редактирования ячейки

`Excel.run` имеет перегрузку, которая принимает объект [Excel.RunOptions](/javascript/api/excel/excel.runoptions) . Он содержит набор свойств, влияющих на поведение платформы при выполнении функции. В настоящее время поддерживается следующее свойство.

- `delayForCellEdit`: определяет, откладывает ли Excel пакетный запрос до выхода пользователя из режима правки ячейки. Если `true`пакетный запрос задерживается и выполняется, когда пользователь выходит из режима редактирования ячеек. При `false`этом пакетный запрос автоматически завершается сбоем, если пользователь находится в режиме редактирования ячейки (что приводит к ошибке для доступа к пользователю). Поведение по умолчанию без указанного `delayForCellEdit` свойства эквивалентно его значению `false`.

```js
await Excel.run({ delayForCellEdit: true }, async (context) => { ... });
```
