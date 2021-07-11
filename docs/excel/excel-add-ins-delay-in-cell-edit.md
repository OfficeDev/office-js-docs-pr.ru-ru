---
title: Задержка выполнения во время редактирования ячейки
description: Узнайте, как отсрочить выполнение метода Excel.run при редактировании ячейки.
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: b7b28064ef4d313639391e63cba780351b5623f9
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349520"
---
# <a name="delay-execution-while-cell-is-being-edited"></a><span data-ttu-id="91ee9-103">Задержка выполнения во время редактирования ячейки</span><span class="sxs-lookup"><span data-stu-id="91ee9-103">Delay execution while cell is being edited</span></span>

<span data-ttu-id="91ee9-104">`Excel.run`имеет перегрузку, которая принимает в [Excel. Объект RunOptions.](/javascript/api/excel/excel.runoptions)</span><span class="sxs-lookup"><span data-stu-id="91ee9-104">`Excel.run` has an overload that takes in a [Excel.RunOptions](/javascript/api/excel/excel.runoptions) object.</span></span> <span data-ttu-id="91ee9-105">Он содержит набор свойств, влияющих на поведение платформы при выполнении функции.</span><span class="sxs-lookup"><span data-stu-id="91ee9-105">This contains a set of properties that affect platform behavior when the function runs.</span></span> <span data-ttu-id="91ee9-106">В настоящее время поддерживается следующее свойство.</span><span class="sxs-lookup"><span data-stu-id="91ee9-106">The following property is currently supported.</span></span>

- <span data-ttu-id="91ee9-107">`delayForCellEdit`: определяет, откладывает ли Excel пакетный запрос до выхода пользователя из режима правки ячейки.</span><span class="sxs-lookup"><span data-stu-id="91ee9-107">`delayForCellEdit`: Determines whether Excel delays the batch request until the user exits cell edit mode.</span></span> <span data-ttu-id="91ee9-108">Если присвоено значение **true**, пакетный запрос откладывается и запускается, когда пользователь выходит из режима правки ячейки.</span><span class="sxs-lookup"><span data-stu-id="91ee9-108">When **true**, the batch request is delayed and runs when the user exits cell edit mode.</span></span> <span data-ttu-id="91ee9-109">Если присвоено значение **false**, происходит автоматический сбой пакетного запроса, если пользователь находится в режиме правки ячейки (приводит к ошибке обращения к пользователю).</span><span class="sxs-lookup"><span data-stu-id="91ee9-109">When **false**, the batch request automatically fails if the user is in cell edit mode (causing an error to reach the user).</span></span> <span data-ttu-id="91ee9-110">Поведение по умолчанию при отсутствии заданного свойства `delayForCellEdit` аналогично поведению при значении **false**.</span><span class="sxs-lookup"><span data-stu-id="91ee9-110">The default behavior with no `delayForCellEdit` property specified is equivalent to when it is **false**.</span></span>

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```
