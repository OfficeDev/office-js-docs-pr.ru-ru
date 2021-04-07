---
title: Элемент IconUrl в файле манифеста
description: Элемент IconUrl указывает URL-адрес изображения, который представляет вашу надстройку Office в UX и Office Store.
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: 68a449b40f6084d26140d59fec61967e163196df
ms.sourcegitcommit: 0bff0411d8cfefd4bb00c189643358e6fb1df95e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/07/2021
ms.locfileid: "51604641"
---
# <a name="iconurl-element"></a><span data-ttu-id="0d947-103">Элемент IconUrl</span><span class="sxs-lookup"><span data-stu-id="0d947-103">IconUrl element</span></span>

<span data-ttu-id="0d947-104">Указывает URL-адрес изображения, которое используется для представления надстройки Office в пользовательском интерфейсе вставки и Магазине Office.</span><span class="sxs-lookup"><span data-stu-id="0d947-104">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store.</span></span>

<span data-ttu-id="0d947-105">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="0d947-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="0d947-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="0d947-106">Syntax</span></span>

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="0d947-107">Может содержать:</span><span class="sxs-lookup"><span data-stu-id="0d947-107">Can contain</span></span>

[<span data-ttu-id="0d947-108">Override</span><span class="sxs-lookup"><span data-stu-id="0d947-108">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="0d947-109">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0d947-109">Attributes</span></span>

|<span data-ttu-id="0d947-110">Атрибут</span><span class="sxs-lookup"><span data-stu-id="0d947-110">Attribute</span></span>|<span data-ttu-id="0d947-111">Тип</span><span class="sxs-lookup"><span data-stu-id="0d947-111">Type</span></span>|<span data-ttu-id="0d947-112">Обязательный</span><span class="sxs-lookup"><span data-stu-id="0d947-112">Required</span></span>|<span data-ttu-id="0d947-113">Описание</span><span class="sxs-lookup"><span data-stu-id="0d947-113">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="0d947-114">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="0d947-114">DefaultValue</span></span>|<span data-ttu-id="0d947-115">string</span><span class="sxs-lookup"><span data-stu-id="0d947-115">string</span></span>|<span data-ttu-id="0d947-116">Обязательный</span><span class="sxs-lookup"><span data-stu-id="0d947-116">required</span></span>|<span data-ttu-id="0d947-117">Задает значение по умолчанию для этого параметра, представленное для языкового стандарта, который указан с помощью элемента [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="0d947-117">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="0d947-118">Замечания</span><span class="sxs-lookup"><span data-stu-id="0d947-118">Remarks</span></span>

<span data-ttu-id="0d947-119">Для надстройки почты значок отображается в пользовательском интерфейсе управления надстройкой **File**  >  **Manage** (Outlook) или **Settings**  >  **Manage add-ins** UI (Outlook on the web).</span><span class="sxs-lookup"><span data-stu-id="0d947-119">For a mail add-in, the icon is displayed in the **File** > **Manage add-ins** UI (Outlook) or **Settings** > **Manage add-ins** UI (Outlook on the web).</span></span> <span data-ttu-id="0d947-120">Значок надстройки области задач или контентной надстройки отображается в разделе **Вставка** > **Надстройки**.</span><span class="sxs-lookup"><span data-stu-id="0d947-120">For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span> <span data-ttu-id="0d947-121">Для всех типов надстройки значок также используется в [AppSource,](https://appsource.microsoft.com)если вы публикуете надстройку в AppSource.</span><span class="sxs-lookup"><span data-stu-id="0d947-121">For all add-in types, the icon is also used in [AppSource](https://appsource.microsoft.com), if you publish your add-in to AppSource.</span></span>

<span data-ttu-id="0d947-122">Изображение должно быть в формате GIF, JPG, PNG, EXIF, BMP или TIFF.</span><span class="sxs-lookup"><span data-stu-id="0d947-122">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="0d947-123">Для приложений области задач и приложений для работы с контентом указанное изображение должно иметь размеры 32 х 32 пикселя.</span><span class="sxs-lookup"><span data-stu-id="0d947-123">For content and task pane apps, the image specified must be 32 x 32 pixels.</span></span> <span data-ttu-id="0d947-124">Для почтовых приложений разрешение изображения должно быть 64 x 64 пикселя.</span><span class="sxs-lookup"><span data-stu-id="0d947-124">For mail apps, the image resolution must be 64 x 64 pixels.</span></span> <span data-ttu-id="0d947-125">Кроме того, необходимо указать значок для использования в клиентских приложениях Office, работающих на экранах с высоким уровнем DPI с помощью элемента [HighResolutionIconUrl.](highresolutioniconurl.md)</span><span class="sxs-lookup"><span data-stu-id="0d947-125">You should also specify an icon for use with Office client applications running on high DPI screens using the [HighResolutionIconUrl](highresolutioniconurl.md) element.</span></span> <span data-ttu-id="0d947-126">Дополнительные сведения см. в разделе _Создание согласованного визуального образа приложения_ статьи [Создание эффективных описаний в AppSource и в Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span><span class="sxs-lookup"><span data-stu-id="0d947-126">For more information, see the section _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>

<span data-ttu-id="0d947-127">Изменение значения элемента во время запуска в настоящее время `IconUrl` не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="0d947-127">Changing the value of the `IconUrl` element at runtime is not currently supported.</span></span>
