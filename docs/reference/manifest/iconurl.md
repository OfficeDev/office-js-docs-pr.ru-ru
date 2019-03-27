---
title: Элемент IconUrl в файле манифеста
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: f7eda7ec9e4c5da8ad0b19e5e10649696d4e85c1
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871726"
---
# <a name="iconurl-element"></a><span data-ttu-id="1909a-102">Элемент IconUrl</span><span class="sxs-lookup"><span data-stu-id="1909a-102">IconUrl element</span></span>

<span data-ttu-id="1909a-103">Указывает URL-адрес изображения, которое используется для представления надстройки Office в пользовательском интерфейсе вставки и Магазине Office.</span><span class="sxs-lookup"><span data-stu-id="1909a-103">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store.</span></span>

<span data-ttu-id="1909a-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="1909a-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="1909a-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="1909a-105">Syntax</span></span>

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="1909a-106">Может содержать:</span><span class="sxs-lookup"><span data-stu-id="1909a-106">Can contain</span></span>

[<span data-ttu-id="1909a-107">Override</span><span class="sxs-lookup"><span data-stu-id="1909a-107">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="1909a-108">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="1909a-108">Attributes</span></span>

|<span data-ttu-id="1909a-109">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="1909a-109">**Attribute**</span></span>|<span data-ttu-id="1909a-110">**Тип**</span><span class="sxs-lookup"><span data-stu-id="1909a-110">**Type**</span></span>|<span data-ttu-id="1909a-111">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="1909a-111">**Required**</span></span>|<span data-ttu-id="1909a-112">**Описание**</span><span class="sxs-lookup"><span data-stu-id="1909a-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="1909a-113">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="1909a-113">DefaultValue</span></span>|<span data-ttu-id="1909a-114">string</span><span class="sxs-lookup"><span data-stu-id="1909a-114">string</span></span>|<span data-ttu-id="1909a-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="1909a-115">required</span></span>|<span data-ttu-id="1909a-116">Задает значение по умолчанию для этого параметра, представленное для языкового стандарта, который указан с помощью элемента [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="1909a-116">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="1909a-117">Замечания</span><span class="sxs-lookup"><span data-stu-id="1909a-117">Remarks</span></span>

<span data-ttu-id="1909a-p101">Значок почтовой надстройки отображается в разделе **Файл**  >  **Управление надстройками** (Outlook) или **Параметры**  >  **Управление надстройками** UI (Outlook Web App). Значок надстройки области задач или контентной надстройки отображается в разделе **Вставка**  >  **Надстройки**. В случае всех типов надстроек значок также используется на сайте Магазина Office, если надстройка опубликована там.</span><span class="sxs-lookup"><span data-stu-id="1909a-p101">For a mail add-in, the icon is displayed in the  **File** > **Manage add-ins** UI (Outlook) or **Settings** > **Manage add-ins** UI (Outlook Web App). For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI. For all add-in types, the icon is also used on the Office Store site, if you publish your add-in to the Office Store.</span></span>

<span data-ttu-id="1909a-121">Изображение должно быть в формате GIF, JPG, PNG, EXIF, BMP или TIFF.</span><span class="sxs-lookup"><span data-stu-id="1909a-121">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="1909a-122">Для приложений области задач и приложений для работы с контентом указанное изображение должно иметь размеры 32 х 32 пикселя.</span><span class="sxs-lookup"><span data-stu-id="1909a-122">For content and task pane apps, the image specified must be 32 x 32 pixels.</span></span> <span data-ttu-id="1909a-123">Для почтовых приложений рекомендуется размер изображения 64 х 64 пикселя.</span><span class="sxs-lookup"><span data-stu-id="1909a-123">For mail apps, the recommended image resolution is 64 x 64 pixels.</span></span> <span data-ttu-id="1909a-124">Кроме того, следует указать значок, который будет использоваться в ведущих приложениях Office на экранах c высоким DPI, при помощи элемента [HighResolutionIconUrl](highresolutioniconurl.md).</span><span class="sxs-lookup"><span data-stu-id="1909a-124">You should also specify an icon for use with Office host applications running on high DPI screens using the [HighResolutionIconUrl](highresolutioniconurl.md) element.</span></span> <span data-ttu-id="1909a-125">Дополнительные сведения см. в разделе _Создание согласованного визуального образа приложения_ статьи [Создание эффективных описаний в AppSource и в Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span><span class="sxs-lookup"><span data-stu-id="1909a-125">For more information, see the section _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>
