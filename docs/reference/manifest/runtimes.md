---
title: Среды выполнения в файле манифеста (Предварительная версия)
description: Элемент Runtimes указывает среду выполнения надстройки.
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 5797aa78ae3667461de48de481ff44f14c307ced
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720423"
---
# <a name="runtimes-element-preview"></a><span data-ttu-id="3580a-103">Элемент среды выполнения (Предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="3580a-103">Runtimes element (preview)</span></span>

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="3580a-104">Указывает среду выполнения надстройки и позволяет использовать пользовательские функции, кнопки ленты и область задач для использования одной и той же среды выполнения JavaScript.</span><span class="sxs-lookup"><span data-stu-id="3580a-104">Specifies the runtime of your add-in and enables custom functions, ribbon buttons, and the task pane to use the same JavaScript runtime.</span></span> <span data-ttu-id="3580a-105">Дочерний `<Host>` элемент элемента в файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="3580a-105">Child of the `<Host>` element in your manifest file.</span></span> <span data-ttu-id="3580a-106">Дополнительные сведения можно найти в статье [Настройка надстройки Excel для использования общей среды выполнения JavaScript](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="3580a-106">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="3580a-107">**Тип надстройки:** надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="3580a-107">**Add-in type:** Task pane</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3580a-108">Общедоступная среда выполнения в настоящее время находится в режиме предварительной версии и доступна только в Excel для Windows.</span><span class="sxs-lookup"><span data-stu-id="3580a-108">Shared runtime is currently in preview and are only available on Excel on Windows.</span></span> <span data-ttu-id="3580a-109">Для ознакомления с предварительными возможностями необходимо присоединиться к [программе предварительной оценки Office](https://insider.office.com/).</span><span class="sxs-lookup"><span data-stu-id="3580a-109">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>

## <a name="syntax"></a><span data-ttu-id="3580a-110">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="3580a-110">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="3580a-111">Содержится в</span><span class="sxs-lookup"><span data-stu-id="3580a-111">Contained in</span></span> 
[<span data-ttu-id="3580a-112">Host</span><span class="sxs-lookup"><span data-stu-id="3580a-112">Host</span></span>](./host.md)

## <a name="child-elements"></a><span data-ttu-id="3580a-113">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="3580a-113">Child elements</span></span>

|  <span data-ttu-id="3580a-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="3580a-114">Element</span></span> |  <span data-ttu-id="3580a-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="3580a-115">Required</span></span>  |  <span data-ttu-id="3580a-116">Описание</span><span class="sxs-lookup"><span data-stu-id="3580a-116">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="3580a-117">**Среда выполнения**</span><span class="sxs-lookup"><span data-stu-id="3580a-117">**Runtime**</span></span>     | <span data-ttu-id="3580a-118">Да</span><span class="sxs-lookup"><span data-stu-id="3580a-118">Yes</span></span> |  <span data-ttu-id="3580a-119">Среда выполнения надстройки.</span><span class="sxs-lookup"><span data-stu-id="3580a-119">The runtime for your add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="3580a-120">См. также</span><span class="sxs-lookup"><span data-stu-id="3580a-120">See also</span></span>

- [<span data-ttu-id="3580a-121">Среда выполнения</span><span class="sxs-lookup"><span data-stu-id="3580a-121">Runtime</span></span>](runtime.md)
