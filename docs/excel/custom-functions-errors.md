---
ms.date: 02/08/2019
description: Обработка ошибок в пользовательских функциях Excel.
title: Обработка ошибок в пользовательских функциях Excel (предварительная версия)
localization_priority: Priority
ms.openlocfilehash: 6c1c7f780aea125977510e4eb0e320933cd6ed9c
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870011"
---
# <a name="error-handling-within-custom-functions"></a><span data-ttu-id="f6ee3-103">Обработка ошибок в пользовательских функциях</span><span class="sxs-lookup"><span data-stu-id="f6ee3-103">Error handling within custom functions</span></span>

<span data-ttu-id="f6ee3-104">При создании надстройки, которая определяет пользовательские функции, не забудьте включить логику для обработки ошибок, возникающих в среде выполнения.</span><span class="sxs-lookup"><span data-stu-id="f6ee3-104">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="f6ee3-105">Обработка ошибок для пользовательских функций в значительной степени совпадает с [обработкой ошибок для API JavaScript в Excel](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="f6ee3-105">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="f6ee3-106">В следующем примере кода `.catch` будет обрабатывать любые ошибки, возникающие ранее в коде.</span><span class="sxs-lookup"><span data-stu-id="f6ee3-106">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

```js
function getComment(commentID) {
  let url = "https://www.contoso.com/comments/" + x;

  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then(function (json) {
      return json.body;
    })
    .catch(function (error) {
      throw error;
    })
}
```

## <a name="see-also"></a><span data-ttu-id="f6ee3-107">См. также</span><span class="sxs-lookup"><span data-stu-id="f6ee3-107">See also</span></span>

* [<span data-ttu-id="f6ee3-108">Руководство по пользовательским функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="f6ee3-108">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="f6ee3-109">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="f6ee3-109">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="f6ee3-110">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="f6ee3-110">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="f6ee3-111">Рекомендации по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="f6ee3-111">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="f6ee3-112">Журнал изменений пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="f6ee3-112">Custom functions changelog</span></span>](custom-functions-changelog.md)
