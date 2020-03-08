---
title: Среда выполнения в файле манифеста (Предварительная версия)
description: ''
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: dd51c5b317700f92ee74c94835e68523371789f8
ms.sourcegitcommit: 153576b1efd0234c6252433e22db213238573534
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/07/2020
ms.locfileid: "42561830"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="8b0a2-102">Элемент среды выполнения (Предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="8b0a2-102">Runtime element (preview)</span></span>

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="8b0a2-103">Дочерний элемент [`<Runtimes>`](runtimes.md) элемента.</span><span class="sxs-lookup"><span data-stu-id="8b0a2-103">Child element of the [`<Runtimes>`](runtimes.md) element.</span></span> <span data-ttu-id="8b0a2-104">Этот элемент настраивает надстройку, чтобы использовать общую среду выполнения JavaScript, чтобы Ваша лента, область задач и пользовательские функции выполнялись в одной среде выполнения.</span><span class="sxs-lookup"><span data-stu-id="8b0a2-104">This element configures your add-in to use a shared JavaScript runtime so that your ribbon, task pane, and custom functions, all run in the same runtime.</span></span> <span data-ttu-id="8b0a2-105">Дополнительные сведения можно найти в статье [Настройка надстройки Excel для использования общей среды выполнения JavaScript](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="8b0a2-105">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="8b0a2-106">**Тип надстройки:** надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="8b0a2-106">**Add-in type:** Task pane</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8b0a2-107">Общедоступная среда выполнения в настоящее время находится в режиме предварительной версии и доступна только в Excel для Windows.</span><span class="sxs-lookup"><span data-stu-id="8b0a2-107">Shared runtime is currently in preview and are only available on Excel on Windows.</span></span> <span data-ttu-id="8b0a2-108">Для ознакомления с предварительными возможностями необходимо присоединиться к [программе предварительной оценки Office](https://insider.office.com/).</span><span class="sxs-lookup"><span data-stu-id="8b0a2-108">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>

## <a name="syntax"></a><span data-ttu-id="8b0a2-109">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="8b0a2-109">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="8b0a2-110">Содержится в</span><span class="sxs-lookup"><span data-stu-id="8b0a2-110">Contained in</span></span>

- [<span data-ttu-id="8b0a2-111">Runtimes</span><span class="sxs-lookup"><span data-stu-id="8b0a2-111">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="8b0a2-112">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="8b0a2-112">Attributes</span></span>

|  <span data-ttu-id="8b0a2-113">Атрибут</span><span class="sxs-lookup"><span data-stu-id="8b0a2-113">Attribute</span></span>  |  <span data-ttu-id="8b0a2-114">Обязательный</span><span class="sxs-lookup"><span data-stu-id="8b0a2-114">Required</span></span>  |  <span data-ttu-id="8b0a2-115">Описание</span><span class="sxs-lookup"><span data-stu-id="8b0a2-115">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="8b0a2-116">**время жизни = "Long"**</span><span class="sxs-lookup"><span data-stu-id="8b0a2-116">**lifetime="long"**</span></span>  |  <span data-ttu-id="8b0a2-117">Да</span><span class="sxs-lookup"><span data-stu-id="8b0a2-117">Yes</span></span>  | <span data-ttu-id="8b0a2-118">Всегда следует использовать `long` , если вы хотите использовать общую среду выполнения для надстройки Excel.</span><span class="sxs-lookup"><span data-stu-id="8b0a2-118">Should always be `long` if you want to use a shared runtime for the Excel add-in.</span></span> |
|  <span data-ttu-id="8b0a2-119">**resid**</span><span class="sxs-lookup"><span data-stu-id="8b0a2-119">**resid**</span></span>  |  <span data-ttu-id="8b0a2-120">Да</span><span class="sxs-lookup"><span data-stu-id="8b0a2-120">Yes</span></span>  | <span data-ttu-id="8b0a2-121">Указывает URL-адрес HTML-страницы для надстройки.</span><span class="sxs-lookup"><span data-stu-id="8b0a2-121">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="8b0a2-122">`resid` Должен сопоставляться с `id` атрибутом `Url` элемента в `Resources` элементе.</span><span class="sxs-lookup"><span data-stu-id="8b0a2-122">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |

## <a name="see-also"></a><span data-ttu-id="8b0a2-123">См. также</span><span class="sxs-lookup"><span data-stu-id="8b0a2-123">See also</span></span>

- [<span data-ttu-id="8b0a2-124">Runtimes</span><span class="sxs-lookup"><span data-stu-id="8b0a2-124">Runtimes</span></span>](runtimes.md)
