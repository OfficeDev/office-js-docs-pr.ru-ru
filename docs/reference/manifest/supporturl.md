---
title: Элемент SupportUrl в файле манифеста
description: Элемент SupportUrl указывает URL-адрес страницы, предоставляющей сведения о поддержке надстройки.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: e38030062c48936f925126e896cd74e660164a5d
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720346"
---
# <a name="supporturl-element"></a><span data-ttu-id="3e05a-103">Элемент SupportUrl</span><span class="sxs-lookup"><span data-stu-id="3e05a-103">SupportUrl element</span></span>

<span data-ttu-id="3e05a-104">Указывает URL-адрес страницы, на которой представлены сведения о поддержке надстройки.</span><span class="sxs-lookup"><span data-stu-id="3e05a-104">Specifies the URL of a page that provides support information for your add-in.</span></span>

## <a name="syntax"></a><span data-ttu-id="3e05a-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="3e05a-105">Syntax</span></span>

```XML
<OfficeApp>
...
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png"/>
  
  
  <SupportUrl DefaultValue="https://contoso.com/support " />
  
  
  <AppDomains>
  ...
  </AppDomains>
...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="3e05a-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="3e05a-106">Contained in</span></span>

[<span data-ttu-id="3e05a-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="3e05a-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="3e05a-108">Может содержать</span><span class="sxs-lookup"><span data-stu-id="3e05a-108">Can contain</span></span>

|  <span data-ttu-id="3e05a-109">Элемент</span><span class="sxs-lookup"><span data-stu-id="3e05a-109">Element</span></span> | <span data-ttu-id="3e05a-110">Обязательный</span><span class="sxs-lookup"><span data-stu-id="3e05a-110">Required</span></span> | <span data-ttu-id="3e05a-111">Описание</span><span class="sxs-lookup"><span data-stu-id="3e05a-111">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="3e05a-112">Override</span><span class="sxs-lookup"><span data-stu-id="3e05a-112">Override</span></span>](override.md)   | <span data-ttu-id="3e05a-113">Нет</span><span class="sxs-lookup"><span data-stu-id="3e05a-113">No</span></span> | <span data-ttu-id="3e05a-114">Задает параметр для URL-адресов дополнительных языковых стандартов</span><span class="sxs-lookup"><span data-stu-id="3e05a-114">Specifies the setting for additional locale urls</span></span> |

## <a name="attributes"></a><span data-ttu-id="3e05a-115">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="3e05a-115">Attributes</span></span>

|<span data-ttu-id="3e05a-116">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="3e05a-116">**Attribute**</span></span>|<span data-ttu-id="3e05a-117">**Тип**</span><span class="sxs-lookup"><span data-stu-id="3e05a-117">**Type**</span></span>|<span data-ttu-id="3e05a-118">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="3e05a-118">**Required**</span></span>|<span data-ttu-id="3e05a-119">**Описание**</span><span class="sxs-lookup"><span data-stu-id="3e05a-119">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="3e05a-120">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="3e05a-120">DefaultValue</span></span>|<span data-ttu-id="3e05a-121">URL-адрес</span><span class="sxs-lookup"><span data-stu-id="3e05a-121">URL</span></span>|<span data-ttu-id="3e05a-122">Обязательный</span><span class="sxs-lookup"><span data-stu-id="3e05a-122">required</span></span>|<span data-ttu-id="3e05a-123">Задает значение по умолчанию для этого параметра, представленное для языкового стандарта, который указан с помощью элемента [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="3e05a-123">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
