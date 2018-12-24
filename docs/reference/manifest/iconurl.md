---
title: Элемент IconUrl в файле манифеста
description: ''
ms.date: 12/04/2018
ms.openlocfilehash: 471a168b5aa0091292132a1e078fa2b3f5efb448
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433126"
---
# <a name="iconurl-element"></a><span data-ttu-id="2657c-102">Элемент IconUrl</span><span class="sxs-lookup"><span data-stu-id="2657c-102">IconUrl element</span></span>

<span data-ttu-id="2657c-103">Указывает URL-адрес изображения, которое используется для представления надстройки Office в пользовательском интерфейсе вставки и Магазине Office.</span><span class="sxs-lookup"><span data-stu-id="2657c-103">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store.</span></span>

<span data-ttu-id="2657c-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="2657c-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="2657c-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="2657c-105">Syntax</span></span>

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="2657c-106">Может содержать:</span><span class="sxs-lookup"><span data-stu-id="2657c-106">Can contain</span></span>

[<span data-ttu-id="2657c-107">Переопределение</span><span class="sxs-lookup"><span data-stu-id="2657c-107">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="2657c-108">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="2657c-108">Attributes</span></span>

|<span data-ttu-id="2657c-109">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="2657c-109">**Attribute**</span></span>|<span data-ttu-id="2657c-110">**Тип**</span><span class="sxs-lookup"><span data-stu-id="2657c-110">**Type**</span></span>|<span data-ttu-id="2657c-111">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="2657c-111">**Required**</span></span>|<span data-ttu-id="2657c-112">**Описание**</span><span class="sxs-lookup"><span data-stu-id="2657c-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="2657c-113">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="2657c-113">DefaultValue</span></span>|<span data-ttu-id="2657c-114">string</span><span class="sxs-lookup"><span data-stu-id="2657c-114">string</span></span>|<span data-ttu-id="2657c-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="2657c-115">required</span></span>|<span data-ttu-id="2657c-116">Задает значение по умолчанию для этого параметра, представленное для языкового стандарта, который указан с помощью элемента [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="2657c-116">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="2657c-117">Замечания</span><span class="sxs-lookup"><span data-stu-id="2657c-117">Remarks</span></span>

<span data-ttu-id="2657c-p101">Значок почтовой надстройки отображается в разделе **Файл**  >  **Управление надстройками** (Outlook) или **Параметры**  >  **Управление надстройками** UI (Outlook Web App). Значок надстройки области задач или контентной надстройки отображается в разделе **Вставка**  >  **Надстройки**. В случае всех типов надстроек значок также используется на сайте Магазина Office, если надстройка опубликована там.</span><span class="sxs-lookup"><span data-stu-id="2657c-p101">For a mail add-in, the icon is displayed in the  **File** > **Manage add-ins** UI (Outlook) or **Settings** > **Manage add-ins** UI (Outlook Web App). For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI. For all add-in types, the icon is also used on the Office Store site, if you publish your add-in to the Office Store.</span></span>

<span data-ttu-id="2657c-121">Изображение должно быть в формате GIF, JPG, PNG, EXIF, BMP или TIFF.</span><span class="sxs-lookup"><span data-stu-id="2657c-121">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="2657c-122">Для приложений области задач и приложений для работы с контентом указанное изображение должно иметь размеры 32 х 32 пикселя.</span><span class="sxs-lookup"><span data-stu-id="2657c-122">For content and task pane apps, the image specified must be 32 x 32 pixels.</span></span> <span data-ttu-id="2657c-123">Для почтовых приложений рекомендуется размер изображения 64 х 64 пикселя.</span><span class="sxs-lookup"><span data-stu-id="2657c-123">For mail apps, the recommended image resolution is 64 x 64 pixels.</span></span> <span data-ttu-id="2657c-124">Кроме того, следует указать значок, который будет использоваться в ведущих приложениях Office на экранах c высоким DPI, при помощи элемента [HighResolutionIconUrl](highresolutioniconurl.md).</span><span class="sxs-lookup"><span data-stu-id="2657c-124">You should also specify an icon for use with Office host applications running on high DPI screens using the [HighResolutionIconUrl](highresolutioniconurl.md) element.</span></span> <span data-ttu-id="2657c-125">Дополнительные сведения см. в разделе _Создание согласованного визуального образа приложения_ статьи [Создание эффективных описаний в AppSource и в Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span><span class="sxs-lookup"><span data-stu-id="2657c-125">For more information, see the section _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>
