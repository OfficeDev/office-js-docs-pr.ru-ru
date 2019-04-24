---
title: Элемент SupportUrl в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 18b9b7c4df9def70ab42ae213066188ac04c07a7
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450417"
---
# <a name="supporturl-element"></a><span data-ttu-id="862ec-102">Элемент SupportUrl</span><span class="sxs-lookup"><span data-stu-id="862ec-102">SupportUrl element</span></span>

<span data-ttu-id="862ec-103">Указывает URL-адрес страницы, на которой представлены сведения о поддержке надстройки.</span><span class="sxs-lookup"><span data-stu-id="862ec-103">Specifies the URL of a page that provides support information for your add-in.</span></span>

## <a name="syntax"></a><span data-ttu-id="862ec-104">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="862ec-104">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="862ec-105">Содержится в</span><span class="sxs-lookup"><span data-stu-id="862ec-105">Contained in</span></span>

[<span data-ttu-id="862ec-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="862ec-106">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="862ec-107">Может содержать</span><span class="sxs-lookup"><span data-stu-id="862ec-107">Can contain</span></span>

|  <span data-ttu-id="862ec-108">Элемент</span><span class="sxs-lookup"><span data-stu-id="862ec-108">Element</span></span> | <span data-ttu-id="862ec-109">Обязательный</span><span class="sxs-lookup"><span data-stu-id="862ec-109">Required</span></span> | <span data-ttu-id="862ec-110">Описание</span><span class="sxs-lookup"><span data-stu-id="862ec-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="862ec-111">Override</span><span class="sxs-lookup"><span data-stu-id="862ec-111">Override</span></span>](override.md)   | <span data-ttu-id="862ec-112">Нет</span><span class="sxs-lookup"><span data-stu-id="862ec-112">No</span></span> | <span data-ttu-id="862ec-113">Задает параметр для URL-адресов дополнительных языковых стандартов</span><span class="sxs-lookup"><span data-stu-id="862ec-113">Specifies the setting for additional locale urls</span></span> |

## <a name="attributes"></a><span data-ttu-id="862ec-114">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="862ec-114">Attributes</span></span>

|<span data-ttu-id="862ec-115">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="862ec-115">**Attribute**</span></span>|<span data-ttu-id="862ec-116">**Тип**</span><span class="sxs-lookup"><span data-stu-id="862ec-116">**Type**</span></span>|<span data-ttu-id="862ec-117">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="862ec-117">**Required**</span></span>|<span data-ttu-id="862ec-118">**Описание**</span><span class="sxs-lookup"><span data-stu-id="862ec-118">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="862ec-119">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="862ec-119">DefaultValue</span></span>|<span data-ttu-id="862ec-120">URL-адрес</span><span class="sxs-lookup"><span data-stu-id="862ec-120">URL</span></span>|<span data-ttu-id="862ec-121">Обязательный</span><span class="sxs-lookup"><span data-stu-id="862ec-121">required</span></span>|<span data-ttu-id="862ec-122">Задает значение по умолчанию для этого параметра, представленное для языкового стандарта, который указан с помощью элемента [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="862ec-122">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
