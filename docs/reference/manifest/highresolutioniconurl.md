---
title: Элемент HighResolutionIconUrl в файле манифеста
description: ''
ms.date: 12/04/2018
ms.openlocfilehash: dc8feb92eb8a53351679834a39c012b47f43aad4
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432595"
---
# <a name="highresolutioniconurl-element"></a>Элемент HighResolutionIconUrl

Указывает URL-адрес изображения, которое используется для представления надстройки Office в пользовательском интерфейсе вставки и Магазине Office на экранах с высоким DPI.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.

## <a name="syntax"></a>Синтаксис

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a>Может содержать:

[Переопределение](override.md)

## <a name="attributes"></a>Атрибуты

|**Атрибут**|**Тип**|**Обязательный**|**Описание**|
|:-----|:-----|:-----|:-----|
|DefaultValue|string (URL-адрес)|Обязательный|Задает значение по умолчанию для этого параметра, представленное для языкового стандарта, который указан с помощью элемента [DefaultLocale](defaultlocale.md).|

## <a name="remarks"></a>Замечания

Значок почтовой надстройки отображается в разделе **Файл**  >  **Управление надстройками**. Значок надстройки области задач или контентной надстройки отображается в разделе **Вставка**  >  **Надстройки**.

Изображение должно быть в формате GIF, JPG, PNG, EXIF, BMP или TIFF. Для приложений области задач и приложений для работы с контентом рекомендуется размер изображения 64 х 64 пикселя. Для почтовых приложений изображение должно иметь размер 128 x 128 пикселей. Дополнительные сведения см. в разделе _Создание согласованного визуального образа приложения_ статьи [Создание эффективных описаний в AppSource и в Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).
