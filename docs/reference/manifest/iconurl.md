---
title: Элемент IconUrl в файле манифеста
description: Элемент IconUrl указывает URL-адрес изображения, которое представляет надстройку Office в пользовательском интерфейсе вставки и магазине Office.
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: b2c117f69259911bf4643eaddbbbcdf5fb502197
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641363"
---
# <a name="iconurl-element"></a><span data-ttu-id="67098-103">Элемент IconUrl</span><span class="sxs-lookup"><span data-stu-id="67098-103">IconUrl element</span></span>

<span data-ttu-id="67098-104">Указывает URL-адрес изображения, которое используется для представления надстройки Office в пользовательском интерфейсе вставки и Магазине Office.</span><span class="sxs-lookup"><span data-stu-id="67098-104">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store.</span></span>

<span data-ttu-id="67098-105">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="67098-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="67098-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="67098-106">Syntax</span></span>

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="67098-107">Может содержать:</span><span class="sxs-lookup"><span data-stu-id="67098-107">Can contain</span></span>

[<span data-ttu-id="67098-108">Override</span><span class="sxs-lookup"><span data-stu-id="67098-108">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="67098-109">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="67098-109">Attributes</span></span>

|<span data-ttu-id="67098-110">Атрибут</span><span class="sxs-lookup"><span data-stu-id="67098-110">Attribute</span></span>|<span data-ttu-id="67098-111">Тип</span><span class="sxs-lookup"><span data-stu-id="67098-111">Type</span></span>|<span data-ttu-id="67098-112">Обязательный</span><span class="sxs-lookup"><span data-stu-id="67098-112">Required</span></span>|<span data-ttu-id="67098-113">Описание</span><span class="sxs-lookup"><span data-stu-id="67098-113">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="67098-114">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="67098-114">DefaultValue</span></span>|<span data-ttu-id="67098-115">string</span><span class="sxs-lookup"><span data-stu-id="67098-115">string</span></span>|<span data-ttu-id="67098-116">Обязательный</span><span class="sxs-lookup"><span data-stu-id="67098-116">required</span></span>|<span data-ttu-id="67098-117">Задает значение по умолчанию для этого параметра, представленное для языкового стандарта, который указан с помощью элемента [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="67098-117">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="67098-118">Замечания</span><span class="sxs-lookup"><span data-stu-id="67098-118">Remarks</span></span>

<span data-ttu-id="67098-119">Для почтовой надстройки значок отображается в **File**  >  пользовательском интерфейсе**Управление** файлами надстроек (Outlook) или надстройки **Settings**  >  **управления** параметрами (Outlook в Интернете).</span><span class="sxs-lookup"><span data-stu-id="67098-119">For a mail add-in, the icon is displayed in the **File** > **Manage add-ins** UI (Outlook) or **Settings** > **Manage add-ins** UI (Outlook on the web).</span></span> <span data-ttu-id="67098-120">Значок надстройки области задач или контентной надстройки отображается в разделе **Вставка** > **Надстройки**.</span><span class="sxs-lookup"><span data-stu-id="67098-120">For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span> <span data-ttu-id="67098-121">Для всех типов надстроек значок также используется в [AppSource](https://appsource.microsoft.com), если вы публикуете надстройку в AppSource.</span><span class="sxs-lookup"><span data-stu-id="67098-121">For all add-in types, the icon is also used in [AppSource](https://appsource.microsoft.com), if you publish your add-in to AppSource.</span></span>

<span data-ttu-id="67098-122">Изображение должно быть в формате GIF, JPG, PNG, EXIF, BMP или TIFF.</span><span class="sxs-lookup"><span data-stu-id="67098-122">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="67098-123">Для приложений области задач и приложений для работы с контентом указанное изображение должно иметь размеры 32 х 32 пикселя.</span><span class="sxs-lookup"><span data-stu-id="67098-123">For content and task pane apps, the image specified must be 32 x 32 pixels.</span></span> <span data-ttu-id="67098-124">Для почтовых приложений рекомендуется размер изображения 64 х 64 пикселя.</span><span class="sxs-lookup"><span data-stu-id="67098-124">For mail apps, the recommended image resolution is 64 x 64 pixels.</span></span> <span data-ttu-id="67098-125">Кроме того, следует указать значок, который будет использоваться в ведущих приложениях Office на экранах c высоким DPI, при помощи элемента [HighResolutionIconUrl](highresolutioniconurl.md).</span><span class="sxs-lookup"><span data-stu-id="67098-125">You should also specify an icon for use with Office host applications running on high DPI screens using the [HighResolutionIconUrl](highresolutioniconurl.md) element.</span></span> <span data-ttu-id="67098-126">Дополнительные сведения см. в разделе _Создание согласованного визуального образа приложения_ статьи [Создание эффективных описаний в AppSource и в Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span><span class="sxs-lookup"><span data-stu-id="67098-126">For more information, see the section _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>

<span data-ttu-id="67098-127">Изменение значения `IconUrl` элемента во время выполнения в настоящее время не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="67098-127">Changing the value of the `IconUrl` element at runtime is not currently supported.</span></span>