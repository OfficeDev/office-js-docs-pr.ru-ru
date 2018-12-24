---
title: Элемент SupportUrl в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 00234ef9fe8960b9956e6a2595e2e2e71bfb97c6
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432671"
---
# <a name="supporturl-element"></a><span data-ttu-id="c389c-102">Элемент SupportUrl</span><span class="sxs-lookup"><span data-stu-id="c389c-102">SupportUrl element</span></span>

<span data-ttu-id="c389c-103">Указывает URL-адрес страницы, на которой представлены сведения о поддержке надстройки.</span><span class="sxs-lookup"><span data-stu-id="c389c-103">Specifies the URL of a page that provides support information for your add-in.</span></span>

## <a name="syntax"></a><span data-ttu-id="c389c-104">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="c389c-104">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="c389c-105">Содержится в</span><span class="sxs-lookup"><span data-stu-id="c389c-105">Contained in</span></span>

[<span data-ttu-id="c389c-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="c389c-106">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="c389c-107">Может содержать</span><span class="sxs-lookup"><span data-stu-id="c389c-107">Can contain</span></span>

|  <span data-ttu-id="c389c-108">Элемент</span><span class="sxs-lookup"><span data-stu-id="c389c-108">Element</span></span> | <span data-ttu-id="c389c-109">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c389c-109">Required</span></span> | <span data-ttu-id="c389c-110">Описание</span><span class="sxs-lookup"><span data-stu-id="c389c-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="c389c-111">Override</span><span class="sxs-lookup"><span data-stu-id="c389c-111">Override</span></span>](override.md)   | <span data-ttu-id="c389c-112">Нет</span><span class="sxs-lookup"><span data-stu-id="c389c-112">No</span></span> | <span data-ttu-id="c389c-113">Задает параметр для URL-адресов дополнительных языковых стандартов</span><span class="sxs-lookup"><span data-stu-id="c389c-113">Specifies the setting for additional locale urls</span></span> |

## <a name="attributes"></a><span data-ttu-id="c389c-114">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="c389c-114">Attributes</span></span>

|<span data-ttu-id="c389c-115">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="c389c-115">**Attribute**</span></span>|<span data-ttu-id="c389c-116">**Тип**</span><span class="sxs-lookup"><span data-stu-id="c389c-116">**Type**</span></span>|<span data-ttu-id="c389c-117">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="c389c-117">**Required**</span></span>|<span data-ttu-id="c389c-118">**Описание**</span><span class="sxs-lookup"><span data-stu-id="c389c-118">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="c389c-119">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="c389c-119">DefaultValue</span></span>|<span data-ttu-id="c389c-120">URL-адрес</span><span class="sxs-lookup"><span data-stu-id="c389c-120">URL</span></span>|<span data-ttu-id="c389c-121">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c389c-121">required</span></span>|<span data-ttu-id="c389c-122">Задает значение по умолчанию для этого параметра, представленное для языкового стандарта, который указан с помощью элемента [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="c389c-122">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
