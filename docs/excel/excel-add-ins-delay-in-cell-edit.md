---
title: Задержка выполнения при редактировании ячейки
description: Узнайте, как отложить выполнение метода Excel. Run при редактировании ячейки.
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: eb33f4cb7cce3b1f8642e00f432e708e90b5b895
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2020
ms.locfileid: "47409413"
---
# <a name="delay-execution-while-cell-is-being-edited"></a><span data-ttu-id="c43a0-103">Задержка выполнения при редактировании ячейки</span><span class="sxs-lookup"><span data-stu-id="c43a0-103">Delay execution while cell is being edited</span></span>

<span data-ttu-id="c43a0-104">`Excel.run` имеет перегрузку, которая принимает объект [Excel. руноптионс](/javascript/api/excel/excel.runoptions) .</span><span class="sxs-lookup"><span data-stu-id="c43a0-104">`Excel.run` has an overload that takes in a [Excel.RunOptions](/javascript/api/excel/excel.runoptions) object.</span></span> <span data-ttu-id="c43a0-105">Он содержит набор свойств, влияющих на поведение платформы при выполнении функции.</span><span class="sxs-lookup"><span data-stu-id="c43a0-105">This contains a set of properties that affect platform behavior when the function runs.</span></span> <span data-ttu-id="c43a0-106">Ниже перечислены поддерживаемые в настоящее время свойства.</span><span class="sxs-lookup"><span data-stu-id="c43a0-106">The following property is currently supported:</span></span>

* <span data-ttu-id="c43a0-107">`delayForCellEdit`: определяет, откладывает ли Excel пакетный запрос до выхода пользователя из режима правки ячейки.</span><span class="sxs-lookup"><span data-stu-id="c43a0-107">`delayForCellEdit`: Determines whether Excel delays the batch request until the user exits cell edit mode.</span></span> <span data-ttu-id="c43a0-108">Если присвоено значение **true**, пакетный запрос откладывается и запускается, когда пользователь выходит из режима правки ячейки.</span><span class="sxs-lookup"><span data-stu-id="c43a0-108">When **true**, the batch request is delayed and runs when the user exits cell edit mode.</span></span> <span data-ttu-id="c43a0-109">Если присвоено значение **false**, происходит автоматический сбой пакетного запроса, если пользователь находится в режиме правки ячейки (приводит к ошибке обращения к пользователю).</span><span class="sxs-lookup"><span data-stu-id="c43a0-109">When **false**, the batch request automatically fails if the user is in cell edit mode (causing an error to reach the user).</span></span> <span data-ttu-id="c43a0-110">Поведение по умолчанию при отсутствии заданного свойства `delayForCellEdit` аналогично поведению при значении **false**.</span><span class="sxs-lookup"><span data-stu-id="c43a0-110">The default behavior with no `delayForCellEdit` property specified is equivalent to when it is **false**.</span></span>

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```
