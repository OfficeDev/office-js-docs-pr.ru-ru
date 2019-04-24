---
title: Элемент HighResolutionIconUrl в файле манифеста
description: ''
ms.date: 12/04/2018
localization_priority: Normal
ms.openlocfilehash: 5264fc969bda30a9b2212996800b984533a3188c
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452090"
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

|**Атрибут**|**Тип**|**Обязательный**|**Описание**|
|:-----|:-----|:-----|:-----|
|DefaultValue|string (URL-адрес)|Обязательный|Задает значение по умолчанию для этого параметра, представленное для языкового стандарта, который указан с помощью элемента [DefaultLocale](defaultlocale.md).|

## <a name="remarks"></a>Замечания

Значок почтовой надстройки отображается в разделе **Файл**  >  **Управление надстройками**. Значок надстройки области задач или контентной надстройки отображается в разделе **Вставка**  >  **Надстройки**.

Изображение должно быть в формате GIF, JPG, PNG, EXIF, BMP или TIFF. Для приложений области задач и приложений для работы с контентом рекомендуется размер изображения 64 х 64 пикселя. Для почтовых приложений изображение должно иметь размер 128 x 128 пикселей. Дополнительные сведения см. в разделе _Создание согласованного визуального образа приложения_ статьи [Создание эффективных описаний в AppSource и в Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).
