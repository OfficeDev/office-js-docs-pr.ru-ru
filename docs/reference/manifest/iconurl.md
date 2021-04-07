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
# <a name="iconurl-element"></a>Элемент IconUrl

Указывает URL-адрес изображения, которое используется для представления надстройки Office в пользовательском интерфейсе вставки и Магазине Office.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.

## <a name="syntax"></a>Синтаксис

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a>Может содержать:

[Override](override.md)

## <a name="attributes"></a>Атрибуты

|Атрибут|Тип|Обязательный|Описание|
|:-----|:-----|:-----|:-----|
|DefaultValue|string|Обязательный|Задает значение по умолчанию для этого параметра, представленное для языкового стандарта, который указан с помощью элемента [DefaultLocale](defaultlocale.md).|

## <a name="remarks"></a>Замечания

Для надстройки почты значок отображается в пользовательском интерфейсе управления надстройкой **File**  >  **Manage** (Outlook) или **Settings**  >  **Manage add-ins** UI (Outlook on the web). Значок надстройки области задач или контентной надстройки отображается в разделе **Вставка** > **Надстройки**. Для всех типов надстройки значок также используется в [AppSource,](https://appsource.microsoft.com)если вы публикуете надстройку в AppSource.

Изображение должно быть в формате GIF, JPG, PNG, EXIF, BMP или TIFF. Для приложений области задач и приложений для работы с контентом указанное изображение должно иметь размеры 32 х 32 пикселя. Для почтовых приложений разрешение изображения должно быть 64 x 64 пикселя. Кроме того, необходимо указать значок для использования в клиентских приложениях Office, работающих на экранах с высоким уровнем DPI с помощью элемента [HighResolutionIconUrl.](highresolutioniconurl.md) Дополнительные сведения см. в разделе _Создание согласованного визуального образа приложения_ статьи [Создание эффективных описаний в AppSource и в Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).

Изменение значения элемента во время запуска в настоящее время `IconUrl` не поддерживается.
