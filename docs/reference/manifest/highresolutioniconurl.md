---
title: Элемент HighResolutionIconUrl в файле манифеста
description: Указывает URL-адрес изображения, которое используется для представления надстройки Office в пользовательском интерфейсе вставки и Магазине Office на экранах с высоким DPI.
ms.date: 12/04/2018
localization_priority: Normal
ms.openlocfilehash: 77675e768895a568bdfee97fc4d5006e1e890937
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641356"
---
# <a name="highresolutioniconurl-element"></a><span data-ttu-id="428ae-103">Элемент HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="428ae-103">HighResolutionIconUrl element</span></span>

<span data-ttu-id="428ae-104">Указывает URL-адрес изображения, которое используется для представления надстройки Office в пользовательском интерфейсе вставки и Магазине Office на экранах с высоким DPI.</span><span class="sxs-lookup"><span data-stu-id="428ae-104">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store on high DPI screens.</span></span>

<span data-ttu-id="428ae-105">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="428ae-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="428ae-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="428ae-106">Syntax</span></span>

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="428ae-107">Может содержать:</span><span class="sxs-lookup"><span data-stu-id="428ae-107">Can contain</span></span>

[<span data-ttu-id="428ae-108">Override</span><span class="sxs-lookup"><span data-stu-id="428ae-108">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="428ae-109">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="428ae-109">Attributes</span></span>

|<span data-ttu-id="428ae-110">Атрибут</span><span class="sxs-lookup"><span data-stu-id="428ae-110">Attribute</span></span>|<span data-ttu-id="428ae-111">Тип</span><span class="sxs-lookup"><span data-stu-id="428ae-111">Type</span></span>|<span data-ttu-id="428ae-112">Обязательный</span><span class="sxs-lookup"><span data-stu-id="428ae-112">Required</span></span>|<span data-ttu-id="428ae-113">Описание</span><span class="sxs-lookup"><span data-stu-id="428ae-113">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="428ae-114">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="428ae-114">DefaultValue</span></span>|<span data-ttu-id="428ae-115">string (URL-адрес)</span><span class="sxs-lookup"><span data-stu-id="428ae-115">string (URL)</span></span>|<span data-ttu-id="428ae-116">Обязательный</span><span class="sxs-lookup"><span data-stu-id="428ae-116">required</span></span>|<span data-ttu-id="428ae-117">Задает значение по умолчанию для этого параметра, представленное для языкового стандарта, который указан с помощью элемента [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="428ae-117">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="428ae-118">Замечания</span><span class="sxs-lookup"><span data-stu-id="428ae-118">Remarks</span></span>

<span data-ttu-id="428ae-119">Для почтовой надстройки значок **отображается в**  >  пользовательском интерфейсе**Управление надстройками** .</span><span class="sxs-lookup"><span data-stu-id="428ae-119">For a mail add-in, the icon is displayed in the **File** > **Manage add-ins** UI .</span></span> <span data-ttu-id="428ae-120">Значок надстройки области задач или контентной надстройки отображается в разделе **Вставка** > **Надстройки**.</span><span class="sxs-lookup"><span data-stu-id="428ae-120">For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span>

<span data-ttu-id="428ae-121">Изображение должно быть в формате GIF, JPG, PNG, EXIF, BMP или TIFF.</span><span class="sxs-lookup"><span data-stu-id="428ae-121">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="428ae-122">Для приложений области задач и приложений для работы с контентом рекомендуется размер изображения 64 х 64 пикселя.</span><span class="sxs-lookup"><span data-stu-id="428ae-122">For content and task pane apps, the recommended image resolution is 64 x 64 pixels.</span></span> <span data-ttu-id="428ae-123">Для почтовых приложений изображение должно иметь размер 128 x 128 пикселей.</span><span class="sxs-lookup"><span data-stu-id="428ae-123">For mail apps, the image must be 128 x 128 pixels.</span></span> <span data-ttu-id="428ae-124">Дополнительные сведения см. в разделе _Создание согласованного визуального образа приложения_ статьи [Создание эффективных описаний в AppSource и в Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span><span class="sxs-lookup"><span data-stu-id="428ae-124">For more information, see the section  _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>
