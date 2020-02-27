---
title: Среды выполнения в файле манифеста (Предварительная версия)
description: ''
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 17e53b53d55ea9547cdfc5c4f89f8f4c3a7ab75e
ms.sourcegitcommit: dd6d00202f6466c27418247dad7bd136555a6036
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/26/2020
ms.locfileid: "42283879"
---
# <a name="runtimes-element-preview"></a><span data-ttu-id="eaec5-102">Элемент среды выполнения (Предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="eaec5-102">Runtimes element (preview)</span></span>

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="eaec5-103">Указывает среду выполнения надстройки и позволяет использовать пользовательские функции, кнопки ленты и область задач для использования одной и той же среды выполнения JavaScript.</span><span class="sxs-lookup"><span data-stu-id="eaec5-103">Specifies the runtime of your add-in and enables custom functions, ribbon buttons, and the task pane to use the same JavaScript runtime.</span></span> <span data-ttu-id="eaec5-104">Дочерний `<Host>` элемент элемента в файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="eaec5-104">Child of the `<Host>` element in your manifest file.</span></span> <span data-ttu-id="eaec5-105">Дополнительные сведения можно найти в статье [Настройка надстройки Excel для использования общей среды выполнения JavaScript](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="eaec5-105">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="eaec5-106">**Тип надстройки:** надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="eaec5-106">**Add-in type:** Task pane</span></span>

> [!IMPORTANT]
> <span data-ttu-id="eaec5-107">Общедоступная среда выполнения в настоящее время находится в режиме предварительной версии и доступна только в Excel для Windows.</span><span class="sxs-lookup"><span data-stu-id="eaec5-107">Shared runtime is currently in preview and are only available on Excel on Windows.</span></span> <span data-ttu-id="eaec5-108">Для ознакомления с предварительными возможностями необходимо присоединиться к [программе предварительной оценки Office](https://insider.office.com/).</span><span class="sxs-lookup"><span data-stu-id="eaec5-108">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>

## <a name="syntax"></a><span data-ttu-id="eaec5-109">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="eaec5-109">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="eaec5-110">Содержится в</span><span class="sxs-lookup"><span data-stu-id="eaec5-110">Contained in</span></span> 
[<span data-ttu-id="eaec5-111">Host</span><span class="sxs-lookup"><span data-stu-id="eaec5-111">Host</span></span>](./host.md)

## <a name="child-elements"></a><span data-ttu-id="eaec5-112">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="eaec5-112">Child elements</span></span>

|  <span data-ttu-id="eaec5-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="eaec5-113">Element</span></span> |  <span data-ttu-id="eaec5-114">Обязательный</span><span class="sxs-lookup"><span data-stu-id="eaec5-114">Required</span></span>  |  <span data-ttu-id="eaec5-115">Описание</span><span class="sxs-lookup"><span data-stu-id="eaec5-115">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="eaec5-116">**Среда выполнения**</span><span class="sxs-lookup"><span data-stu-id="eaec5-116">**Runtime**</span></span>     | <span data-ttu-id="eaec5-117">Да</span><span class="sxs-lookup"><span data-stu-id="eaec5-117">Yes</span></span> |  <span data-ttu-id="eaec5-118">Среда выполнения надстройки.</span><span class="sxs-lookup"><span data-stu-id="eaec5-118">The runtime for your add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="eaec5-119">См. также</span><span class="sxs-lookup"><span data-stu-id="eaec5-119">See also</span></span>

- [<span data-ttu-id="eaec5-120">Среда выполнения</span><span class="sxs-lookup"><span data-stu-id="eaec5-120">Runtime</span></span>](runtime.md)
