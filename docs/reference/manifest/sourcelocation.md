---
title: Элемент SourceLocation в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 7544e2bae480b9431c8912533ea1b761132a355e
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451978"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="6dd8a-102">Элемент SourceLocation</span><span class="sxs-lookup"><span data-stu-id="6dd8a-102">SourceLocation element</span></span>

<span data-ttu-id="6dd8a-p101">Указывает расположения исходного файла для надстройки Office как URL-адреса длиной от 1 до 2018 символов. В качестве источника необходимо указать адрес HTTPS, а не путь к файлу.</span><span class="sxs-lookup"><span data-stu-id="6dd8a-p101">Specifies the source file location(s) for your Office Add-in as a URL between 1 and 2018 characters long. The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="6dd8a-105">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="6dd8a-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="6dd8a-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="6dd8a-106">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="6dd8a-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="6dd8a-107">Contained in</span></span>

- <span data-ttu-id="6dd8a-108">[DefaultSettings](defaultsettings.md) (надстройки области задач и контентные надстройки)</span><span class="sxs-lookup"><span data-stu-id="6dd8a-108">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="6dd8a-109">[FormSettings](formsettings.md) (почтовые надстройки)</span><span class="sxs-lookup"><span data-stu-id="6dd8a-109">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="6dd8a-110">[ExtensionPoint](extensionpoint.md) (контекстные почтовые надстройки)</span><span class="sxs-lookup"><span data-stu-id="6dd8a-110">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="6dd8a-111">Может содержать</span><span class="sxs-lookup"><span data-stu-id="6dd8a-111">Can contain</span></span>

[<span data-ttu-id="6dd8a-112">Override</span><span class="sxs-lookup"><span data-stu-id="6dd8a-112">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="6dd8a-113">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="6dd8a-113">Attributes</span></span>

|<span data-ttu-id="6dd8a-114">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="6dd8a-114">**Attribute**</span></span>|<span data-ttu-id="6dd8a-115">**Тип**</span><span class="sxs-lookup"><span data-stu-id="6dd8a-115">**Type**</span></span>|<span data-ttu-id="6dd8a-116">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="6dd8a-116">**Required**</span></span>|<span data-ttu-id="6dd8a-117">**Описание**</span><span class="sxs-lookup"><span data-stu-id="6dd8a-117">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="6dd8a-118">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="6dd8a-118">DefaultValue</span></span>|<span data-ttu-id="6dd8a-119">URL-адрес</span><span class="sxs-lookup"><span data-stu-id="6dd8a-119">URL</span></span>|<span data-ttu-id="6dd8a-120">Обязательный</span><span class="sxs-lookup"><span data-stu-id="6dd8a-120">required</span></span>|<span data-ttu-id="6dd8a-121">Задает значение этого параметра по умолчанию для языкового стандарта, указанного в элементе [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="6dd8a-121">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
