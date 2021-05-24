---
title: Элемент SourceLocation в файле манифеста
description: Элемент SourceLocation указывает расположение исходных файлов для Office надстройки.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 4dcd093db2f23220eaa34c0c81300c4994c1a697
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590899"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="069ef-103">Элемент SourceLocation</span><span class="sxs-lookup"><span data-stu-id="069ef-103">SourceLocation element</span></span>

<span data-ttu-id="069ef-104">Указывает расположение исходных файлов для надстройки Office как URL-адрес длиной от 1 до 2018 символов.</span><span class="sxs-lookup"><span data-stu-id="069ef-104">Specifies the source file locations for your Office Add-in as a URL between 1 and 2018 characters long.</span></span> <span data-ttu-id="069ef-105">В качестве источника необходимо указать адрес HTTPS, а не путь к файлу.</span><span class="sxs-lookup"><span data-stu-id="069ef-105">The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="069ef-106">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="069ef-106">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="069ef-107">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="069ef-107">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="069ef-108">Содержится в</span><span class="sxs-lookup"><span data-stu-id="069ef-108">Contained in</span></span>

- <span data-ttu-id="069ef-109">[DefaultSettings](defaultsettings.md) (надстройки области задач и контентные надстройки)</span><span class="sxs-lookup"><span data-stu-id="069ef-109">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="069ef-110">[FormSettings](formsettings.md) (почтовые надстройки)</span><span class="sxs-lookup"><span data-stu-id="069ef-110">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="069ef-111">[ExtensionPoint](extensionpoint.md) (надстройки для почты Contextual и LaunchEvent)</span><span class="sxs-lookup"><span data-stu-id="069ef-111">[ExtensionPoint](extensionpoint.md) (Contextual and LaunchEvent mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="069ef-112">Может содержать</span><span class="sxs-lookup"><span data-stu-id="069ef-112">Can contain</span></span>

[<span data-ttu-id="069ef-113">Override</span><span class="sxs-lookup"><span data-stu-id="069ef-113">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="069ef-114">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="069ef-114">Attributes</span></span>

|<span data-ttu-id="069ef-115">Атрибут</span><span class="sxs-lookup"><span data-stu-id="069ef-115">Attribute</span></span>|<span data-ttu-id="069ef-116">Тип</span><span class="sxs-lookup"><span data-stu-id="069ef-116">Type</span></span>|<span data-ttu-id="069ef-117">Обязательный</span><span class="sxs-lookup"><span data-stu-id="069ef-117">Required</span></span>|<span data-ttu-id="069ef-118">Описание</span><span class="sxs-lookup"><span data-stu-id="069ef-118">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="069ef-119">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="069ef-119">DefaultValue</span></span>|<span data-ttu-id="069ef-120">URL-адрес</span><span class="sxs-lookup"><span data-stu-id="069ef-120">URL</span></span>|<span data-ttu-id="069ef-121">Обязательный</span><span class="sxs-lookup"><span data-stu-id="069ef-121">required</span></span>|<span data-ttu-id="069ef-122">Задает значение этого параметра по умолчанию для языкового стандарта, указанного в элементе [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="069ef-122">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
