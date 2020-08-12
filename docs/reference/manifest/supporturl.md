---
title: Элемент SupportUrl в файле манифеста
description: Элемент SupportUrl указывает URL-адрес страницы, предоставляющей сведения о поддержке надстройки.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: be516fe5848d775dacb0d424a92be02d59f85512
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641412"
---
# <a name="supporturl-element"></a><span data-ttu-id="32e81-103">Элемент SupportUrl</span><span class="sxs-lookup"><span data-stu-id="32e81-103">SupportUrl element</span></span>

<span data-ttu-id="32e81-104">Указывает URL-адрес страницы, на которой представлены сведения о поддержке надстройки.</span><span class="sxs-lookup"><span data-stu-id="32e81-104">Specifies the URL of a page that provides support information for your add-in.</span></span>

## <a name="syntax"></a><span data-ttu-id="32e81-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="32e81-105">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="32e81-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="32e81-106">Contained in</span></span>

[<span data-ttu-id="32e81-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="32e81-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="32e81-108">Может содержать</span><span class="sxs-lookup"><span data-stu-id="32e81-108">Can contain</span></span>

|  <span data-ttu-id="32e81-109">Элемент</span><span class="sxs-lookup"><span data-stu-id="32e81-109">Element</span></span> | <span data-ttu-id="32e81-110">Обязательный</span><span class="sxs-lookup"><span data-stu-id="32e81-110">Required</span></span> | <span data-ttu-id="32e81-111">Описание</span><span class="sxs-lookup"><span data-stu-id="32e81-111">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="32e81-112">Override</span><span class="sxs-lookup"><span data-stu-id="32e81-112">Override</span></span>](override.md)   | <span data-ttu-id="32e81-113">Нет</span><span class="sxs-lookup"><span data-stu-id="32e81-113">No</span></span> | <span data-ttu-id="32e81-114">Задает параметр для URL-адресов дополнительных языковых стандартов</span><span class="sxs-lookup"><span data-stu-id="32e81-114">Specifies the setting for additional locale urls</span></span> |

## <a name="attributes"></a><span data-ttu-id="32e81-115">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="32e81-115">Attributes</span></span>

|<span data-ttu-id="32e81-116">Атрибут</span><span class="sxs-lookup"><span data-stu-id="32e81-116">Attribute</span></span>|<span data-ttu-id="32e81-117">Тип</span><span class="sxs-lookup"><span data-stu-id="32e81-117">Type</span></span>|<span data-ttu-id="32e81-118">Обязательный</span><span class="sxs-lookup"><span data-stu-id="32e81-118">Required</span></span>|<span data-ttu-id="32e81-119">Описание</span><span class="sxs-lookup"><span data-stu-id="32e81-119">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="32e81-120">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="32e81-120">DefaultValue</span></span>|<span data-ttu-id="32e81-121">URL-адрес</span><span class="sxs-lookup"><span data-stu-id="32e81-121">URL</span></span>|<span data-ttu-id="32e81-122">Обязательный</span><span class="sxs-lookup"><span data-stu-id="32e81-122">required</span></span>|<span data-ttu-id="32e81-123">Задает значение по умолчанию для этого параметра, представленное для языкового стандарта, который указан с помощью элемента [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="32e81-123">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
