---
title: Элемент HighResolutionIconUrl в файле манифеста
description: Указывает URL-адрес изображения, которое используется для представления надстройки Office в пользовательском интерфейсе вставки и Магазине Office на экранах с высоким DPI.
ms.date: 03/30/2021
ms.localizationpriority: medium
ms.openlocfilehash: 83a86b4a27f2cdfea54b657f70400601207464d3
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154782"
---
# <a name="highresolutioniconurl-element"></a>Элемент HighResolutionIconUrl

Указывает URL-адрес изображения, которое используется для представления надстройки Office в пользовательском интерфейсе вставки и Магазине Office на экранах с высоким DPI.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.

## <a name="syntax"></a>Синтаксис

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a>Может содержать:

[Override](override.md)

## <a name="attributes"></a>Атрибуты

|Атрибут|Тип|Обязательный|Описание|
|:-----|:-----|:-----|:-----|
|DefaultValue|string (URL-адрес)|Обязательный|Задает значение по умолчанию для этого параметра, представленное для языкового стандарта, который указан с помощью элемента [DefaultLocale](defaultlocale.md).|

## <a name="remarks"></a>Замечания

Для надстройки почты значок отображается в пользовательском интерфейсе **управления** файлами  >   надстройок. Значок надстройки области задач или контентной надстройки отображается в разделе **Вставка** > **Надстройки**.

Изображение должно быть в формате GIF, JPG, PNG, EXIF, BMP или TIFF. Для приложений для области контента и задач разрешение изображения должно быть 64 x 64 пикселя. Для почтовых приложений изображение должно иметь размер 128 x 128 пикселей. Дополнительные сведения см. в разделе _Создание согласованного визуального образа приложения_ статьи [Создание эффективных описаний в AppSource и в Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).
