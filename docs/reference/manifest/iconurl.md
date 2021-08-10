---
title: Элемент IconUrl в файле манифеста
description: Элемент IconUrl указывает URL-адрес изображения, который Office надстройки в UX и Office Store.
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: c2dac7835dcdd856fb3e713f00b5bd0a3c87189cf36fda3186e51da2c95e1ab9
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57089815"
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

Для надстройки почты значок отображается в пользовательском интерфейсе управления надстройкой File  >  **Manage** (Outlook) или **Параметры**  >  **Manage add-ins** UI (Outlook в Интернете). Значок надстройки области задач или контентной надстройки отображается в разделе **Вставка** > **Надстройки**. Для всех типов надстройки значок также используется в [AppSource,](https://appsource.microsoft.com)если вы публикуете надстройку в AppSource.

Изображение должно быть в формате GIF, JPG, PNG, EXIF, BMP или TIFF. Для приложений области задач и приложений для работы с контентом указанное изображение должно иметь размеры 32 х 32 пикселя. Для почтовых приложений разрешение изображения должно быть 64 x 64 пикселя. Кроме того, необходимо указать значок для Office клиентских приложений, работающих на экранах с высоким уровнем DPI с помощью элемента [HighResolutionIconUrl.](highresolutioniconurl.md) Дополнительные сведения см. в разделе _Создание согласованного визуального образа приложения_ статьи [Создание эффективных описаний в AppSource и в Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).

Изменение значения элемента во время запуска в настоящее время `IconUrl` не поддерживается.
