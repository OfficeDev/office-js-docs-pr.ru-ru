---
title: Среда выполнения в файле манифеста
description: Элемент среды выполнения настраивает надстройку для использования общей среды выполнения JavaScript для ленты, области задач и пользовательских функций.
ms.date: 05/11/2020
localization_priority: Normal
ms.openlocfilehash: c5c7356f9985ca7b5972068629b0587f8916348e
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217762"
---
# <a name="runtime-element"></a><span data-ttu-id="2f070-103">Элемент среды выполнения</span><span class="sxs-lookup"><span data-stu-id="2f070-103">Runtime element</span></span>

<span data-ttu-id="2f070-104">Дочерний элемент [`<Runtimes>`](runtimes.md) элемента.</span><span class="sxs-lookup"><span data-stu-id="2f070-104">Child element of the [`<Runtimes>`](runtimes.md) element.</span></span> <span data-ttu-id="2f070-105">Этот элемент настраивает надстройку, чтобы использовать общую среду выполнения JavaScript, чтобы Ваша лента, область задач и пользовательские функции выполнялись в одной среде выполнения.</span><span class="sxs-lookup"><span data-stu-id="2f070-105">This element configures your add-in to use a shared JavaScript runtime so that your ribbon, task pane, and custom functions, all run in the same runtime.</span></span> <span data-ttu-id="2f070-106">Дополнительные сведения можно найти в статье [Настройка надстройки Excel для использования общей среды выполнения JavaScript](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="2f070-106">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="2f070-107">**Тип надстройки:** надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="2f070-107">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="2f070-108">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="2f070-108">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="2f070-109">Содержится в</span><span class="sxs-lookup"><span data-stu-id="2f070-109">Contained in</span></span>

- [<span data-ttu-id="2f070-110">Runtimes</span><span class="sxs-lookup"><span data-stu-id="2f070-110">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="2f070-111">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="2f070-111">Attributes</span></span>

|  <span data-ttu-id="2f070-112">Атрибут</span><span class="sxs-lookup"><span data-stu-id="2f070-112">Attribute</span></span>  |  <span data-ttu-id="2f070-113">Обязательный</span><span class="sxs-lookup"><span data-stu-id="2f070-113">Required</span></span>  |  <span data-ttu-id="2f070-114">Описание</span><span class="sxs-lookup"><span data-stu-id="2f070-114">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="2f070-115">**время жизни = "Long"**</span><span class="sxs-lookup"><span data-stu-id="2f070-115">**lifetime="long"**</span></span>  |  <span data-ttu-id="2f070-116">Да</span><span class="sxs-lookup"><span data-stu-id="2f070-116">Yes</span></span>  | <span data-ttu-id="2f070-117">Всегда следует `long` использовать, если вы хотите использовать общую среду выполнения для надстройки Excel.</span><span class="sxs-lookup"><span data-stu-id="2f070-117">Should always be `long` if you want to use a shared runtime for the Excel add-in.</span></span> |
|  <span data-ttu-id="2f070-118">**resid**</span><span class="sxs-lookup"><span data-stu-id="2f070-118">**resid**</span></span>  |  <span data-ttu-id="2f070-119">Да</span><span class="sxs-lookup"><span data-stu-id="2f070-119">Yes</span></span>  | <span data-ttu-id="2f070-120">Указывает URL-адрес HTML-страницы для надстройки.</span><span class="sxs-lookup"><span data-stu-id="2f070-120">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="2f070-121">`resid`Должен сопоставляться с `id` атрибутом `Url` элемента в `Resources` элементе.</span><span class="sxs-lookup"><span data-stu-id="2f070-121">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |

## <a name="see-also"></a><span data-ttu-id="2f070-122">См. также</span><span class="sxs-lookup"><span data-stu-id="2f070-122">See also</span></span>

- [<span data-ttu-id="2f070-123">Runtimes</span><span class="sxs-lookup"><span data-stu-id="2f070-123">Runtimes</span></span>](runtimes.md)
