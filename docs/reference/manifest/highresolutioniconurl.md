---
title: Элемент HighResolutionIconUrl в файле манифеста
description: ''
ms.date: 12/04/2018
localization_priority: Normal
ms.openlocfilehash: 41008be6b60d260bef78808af2b8dee1fbd0864a
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325271"
---
# <a name="highresolutioniconurl-element"></a><span data-ttu-id="168f9-102">Элемент HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="168f9-102">HighResolutionIconUrl element</span></span>

<span data-ttu-id="168f9-103">Указывает URL-адрес изображения, которое используется для представления надстройки Office в пользовательском интерфейсе вставки и Магазине Office на экранах с высоким DPI.</span><span class="sxs-lookup"><span data-stu-id="168f9-103">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store on high DPI screens.</span></span>

<span data-ttu-id="168f9-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="168f9-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="168f9-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="168f9-105">Syntax</span></span>

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="168f9-106">Может содержать:</span><span class="sxs-lookup"><span data-stu-id="168f9-106">Can contain</span></span>

[<span data-ttu-id="168f9-107">Override</span><span class="sxs-lookup"><span data-stu-id="168f9-107">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="168f9-108">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="168f9-108">Attributes</span></span>

|<span data-ttu-id="168f9-109">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="168f9-109">**Attribute**</span></span>|<span data-ttu-id="168f9-110">**Тип**</span><span class="sxs-lookup"><span data-stu-id="168f9-110">**Type**</span></span>|<span data-ttu-id="168f9-111">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="168f9-111">**Required**</span></span>|<span data-ttu-id="168f9-112">**Описание**</span><span class="sxs-lookup"><span data-stu-id="168f9-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="168f9-113">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="168f9-113">DefaultValue</span></span>|<span data-ttu-id="168f9-114">string (URL-адрес)</span><span class="sxs-lookup"><span data-stu-id="168f9-114">string (URL)</span></span>|<span data-ttu-id="168f9-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="168f9-115">required</span></span>|<span data-ttu-id="168f9-116">Задает значение по умолчанию для этого параметра, представленное для языкового стандарта, который указан с помощью элемента [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="168f9-116">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="168f9-117">Замечания</span><span class="sxs-lookup"><span data-stu-id="168f9-117">Remarks</span></span>

<span data-ttu-id="168f9-118">Для почтовой надстройки значок **отображается в** > пользовательском интерфейсе**Управление надстройками** .</span><span class="sxs-lookup"><span data-stu-id="168f9-118">For a mail add-in, the icon is displayed in the **File** > **Manage add-ins** UI .</span></span> <span data-ttu-id="168f9-119">Значок надстройки области задач или контентной надстройки отображается в разделе **Вставка** > **Надстройки**.</span><span class="sxs-lookup"><span data-stu-id="168f9-119">For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span>

<span data-ttu-id="168f9-120">Изображение должно быть в формате GIF, JPG, PNG, EXIF, BMP или TIFF.</span><span class="sxs-lookup"><span data-stu-id="168f9-120">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="168f9-121">Для приложений области задач и приложений для работы с контентом рекомендуется размер изображения 64 х 64 пикселя.</span><span class="sxs-lookup"><span data-stu-id="168f9-121">For content and task pane apps, the recommended image resolution is 64 x 64 pixels.</span></span> <span data-ttu-id="168f9-122">Для почтовых приложений изображение должно иметь размер 128 x 128 пикселей.</span><span class="sxs-lookup"><span data-stu-id="168f9-122">For mail apps, the image must be 128 x 128 pixels.</span></span> <span data-ttu-id="168f9-123">Дополнительные сведения см. в разделе _Создание согласованного визуального образа приложения_ статьи [Создание эффективных описаний в AppSource и в Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span><span class="sxs-lookup"><span data-stu-id="168f9-123">For more information, see the section  _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>
