---
title: Среда выполнения в файле манифеста (Предварительная версия)
description: Элемент среды выполнения настраивает надстройку для использования общей среды выполнения JavaScript для ленты, области задач и пользовательских функций.
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 6237f64fec47ed22b0105bf74c8eb7e2b7c38afe
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717931"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="ad47e-103">Элемент среды выполнения (Предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="ad47e-103">Runtime element (preview)</span></span>

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="ad47e-104">Дочерний элемент [`<Runtimes>`](runtimes.md) элемента.</span><span class="sxs-lookup"><span data-stu-id="ad47e-104">Child element of the [`<Runtimes>`](runtimes.md) element.</span></span> <span data-ttu-id="ad47e-105">Этот элемент настраивает надстройку, чтобы использовать общую среду выполнения JavaScript, чтобы Ваша лента, область задач и пользовательские функции выполнялись в одной среде выполнения.</span><span class="sxs-lookup"><span data-stu-id="ad47e-105">This element configures your add-in to use a shared JavaScript runtime so that your ribbon, task pane, and custom functions, all run in the same runtime.</span></span> <span data-ttu-id="ad47e-106">Дополнительные сведения можно найти в статье [Настройка надстройки Excel для использования общей среды выполнения JavaScript](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="ad47e-106">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="ad47e-107">**Тип надстройки:** надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="ad47e-107">**Add-in type:** Task pane</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ad47e-108">Общедоступная среда выполнения в настоящее время находится в режиме предварительной версии и доступна только в Excel для Windows.</span><span class="sxs-lookup"><span data-stu-id="ad47e-108">Shared runtime is currently in preview and are only available on Excel on Windows.</span></span> <span data-ttu-id="ad47e-109">Для ознакомления с предварительными возможностями необходимо присоединиться к [программе предварительной оценки Office](https://insider.office.com/).</span><span class="sxs-lookup"><span data-stu-id="ad47e-109">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>

## <a name="syntax"></a><span data-ttu-id="ad47e-110">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="ad47e-110">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="ad47e-111">Содержится в</span><span class="sxs-lookup"><span data-stu-id="ad47e-111">Contained in</span></span>

- [<span data-ttu-id="ad47e-112">Runtimes</span><span class="sxs-lookup"><span data-stu-id="ad47e-112">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="ad47e-113">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ad47e-113">Attributes</span></span>

|  <span data-ttu-id="ad47e-114">Атрибут</span><span class="sxs-lookup"><span data-stu-id="ad47e-114">Attribute</span></span>  |  <span data-ttu-id="ad47e-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="ad47e-115">Required</span></span>  |  <span data-ttu-id="ad47e-116">Описание</span><span class="sxs-lookup"><span data-stu-id="ad47e-116">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="ad47e-117">**время жизни = "Long"**</span><span class="sxs-lookup"><span data-stu-id="ad47e-117">**lifetime="long"**</span></span>  |  <span data-ttu-id="ad47e-118">Да</span><span class="sxs-lookup"><span data-stu-id="ad47e-118">Yes</span></span>  | <span data-ttu-id="ad47e-119">Всегда следует использовать `long` , если вы хотите использовать общую среду выполнения для надстройки Excel.</span><span class="sxs-lookup"><span data-stu-id="ad47e-119">Should always be `long` if you want to use a shared runtime for the Excel add-in.</span></span> |
|  <span data-ttu-id="ad47e-120">**resid**</span><span class="sxs-lookup"><span data-stu-id="ad47e-120">**resid**</span></span>  |  <span data-ttu-id="ad47e-121">Да</span><span class="sxs-lookup"><span data-stu-id="ad47e-121">Yes</span></span>  | <span data-ttu-id="ad47e-122">Указывает URL-адрес HTML-страницы для надстройки.</span><span class="sxs-lookup"><span data-stu-id="ad47e-122">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="ad47e-123">`resid` Должен сопоставляться с `id` атрибутом `Url` элемента в `Resources` элементе.</span><span class="sxs-lookup"><span data-stu-id="ad47e-123">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |

## <a name="see-also"></a><span data-ttu-id="ad47e-124">См. также</span><span class="sxs-lookup"><span data-stu-id="ad47e-124">See also</span></span>

- [<span data-ttu-id="ad47e-125">Runtimes</span><span class="sxs-lookup"><span data-stu-id="ad47e-125">Runtimes</span></span>](runtimes.md)
