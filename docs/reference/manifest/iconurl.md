---
title: Элемент IconUrl в файле манифеста
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: d4451409a457fa5522e27ab5efd203b9c37a2052
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127557"
---
# <a name="iconurl-element"></a><span data-ttu-id="d5e1f-102">Элемент IconUrl</span><span class="sxs-lookup"><span data-stu-id="d5e1f-102">IconUrl element</span></span>

<span data-ttu-id="d5e1f-103">Указывает URL-адрес изображения, которое используется для представления надстройки Office в пользовательском интерфейсе вставки и Магазине Office.</span><span class="sxs-lookup"><span data-stu-id="d5e1f-103">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store.</span></span>

<span data-ttu-id="d5e1f-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="d5e1f-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="d5e1f-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="d5e1f-105">Syntax</span></span>

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="d5e1f-106">Может содержать:</span><span class="sxs-lookup"><span data-stu-id="d5e1f-106">Can contain</span></span>

[<span data-ttu-id="d5e1f-107">Override</span><span class="sxs-lookup"><span data-stu-id="d5e1f-107">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="d5e1f-108">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d5e1f-108">Attributes</span></span>

|<span data-ttu-id="d5e1f-109">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="d5e1f-109">**Attribute**</span></span>|<span data-ttu-id="d5e1f-110">**Тип**</span><span class="sxs-lookup"><span data-stu-id="d5e1f-110">**Type**</span></span>|<span data-ttu-id="d5e1f-111">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="d5e1f-111">**Required**</span></span>|<span data-ttu-id="d5e1f-112">**Описание**</span><span class="sxs-lookup"><span data-stu-id="d5e1f-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="d5e1f-113">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="d5e1f-113">DefaultValue</span></span>|<span data-ttu-id="d5e1f-114">string</span><span class="sxs-lookup"><span data-stu-id="d5e1f-114">string</span></span>|<span data-ttu-id="d5e1f-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="d5e1f-115">required</span></span>|<span data-ttu-id="d5e1f-116">Задает значение по умолчанию для этого параметра, представленное для языкового стандарта, который указан с помощью элемента [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="d5e1f-116">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="d5e1f-117">Замечания</span><span class="sxs-lookup"><span data-stu-id="d5e1f-117">Remarks</span></span>

<span data-ttu-id="d5e1f-118">Для почтовой надстройки значок отображается в пользовательском интерфейсе**Управление** **файлами** > надстроек (Outlook) или надстройки**управления** **параметрами** > (Outlook в Интернете).</span><span class="sxs-lookup"><span data-stu-id="d5e1f-118">For a mail add-in, the icon is displayed in the  **File** > **Manage add-ins** UI (Outlook) or **Settings** > **Manage add-ins** UI (Outlook on the web).</span></span> <span data-ttu-id="d5e1f-119">Значок надстройки области задач или контентной надстройки отображается в разделе **Вставка** > **Надстройки**.</span><span class="sxs-lookup"><span data-stu-id="d5e1f-119">For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span> <span data-ttu-id="d5e1f-120">В случае всех типов надстроек значок также используется на сайте Магазина Office, если надстройка опубликована там.</span><span class="sxs-lookup"><span data-stu-id="d5e1f-120">For all add-in types, the icon is also used on the Office Store site, if you publish your add-in to the Office Store.</span></span>

<span data-ttu-id="d5e1f-121">Изображение должно быть в формате GIF, JPG, PNG, EXIF, BMP или TIFF.</span><span class="sxs-lookup"><span data-stu-id="d5e1f-121">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="d5e1f-122">Для приложений области задач и приложений для работы с контентом указанное изображение должно иметь размеры 32 х 32 пикселя.</span><span class="sxs-lookup"><span data-stu-id="d5e1f-122">For content and task pane apps, the image specified must be 32 x 32 pixels.</span></span> <span data-ttu-id="d5e1f-123">Для почтовых приложений рекомендуется размер изображения 64 х 64 пикселя.</span><span class="sxs-lookup"><span data-stu-id="d5e1f-123">For mail apps, the recommended image resolution is 64 x 64 pixels.</span></span> <span data-ttu-id="d5e1f-124">Кроме того, следует указать значок, который будет использоваться в ведущих приложениях Office на экранах c высоким DPI, при помощи элемента [HighResolutionIconUrl](highresolutioniconurl.md).</span><span class="sxs-lookup"><span data-stu-id="d5e1f-124">You should also specify an icon for use with Office host applications running on high DPI screens using the [HighResolutionIconUrl](highresolutioniconurl.md) element.</span></span> <span data-ttu-id="d5e1f-125">Дополнительные сведения см. в разделе _Создание согласованного визуального образа приложения_ статьи [Создание эффективных описаний в AppSource и в Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span><span class="sxs-lookup"><span data-stu-id="d5e1f-125">For more information, see the section _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>

<span data-ttu-id="d5e1f-126">Изменение значения `IconUrl` элемента во время выполнения в настоящее время не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="d5e1f-126">Changing the value of the `IconUrl` element at runtime is not currently supported.</span></span>