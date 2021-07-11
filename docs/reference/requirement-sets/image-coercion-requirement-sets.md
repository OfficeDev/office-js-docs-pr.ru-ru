---
title: Наборы обязательных элементов для приведения изображений
description: Поддержка наборов требований к принуждению изображений с Office надстройки в Excel, PowerPoint и Word.
ms.date: 02/19/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 29614718378fd51013360a2a922e11f89bca14b8
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350220"
---
# <a name="image-coercion-requirement-sets"></a>Наборы обязательных элементов для приведения изображений

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

## <a name="imagecoercion-11"></a>ImageCoercion 1.1

ImageCoercion 1.1 позволяет преобразования в изображение () при записи `Office.CoercionType.Image` данных с помощью [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) метода. Поддерживаются следующие приложения.

- Excel 2013 г. и более поздней Windows
- Excel 2016 и позднее на Mac
- Excel на iPad
- OneNote в Интернете
- PowerPoint 2013 и более поздней Windows
- PowerPoint 2016 и более поздней основе на Mac
- PowerPoint в Интернете
- PowerPoint на iPad
- Word 2013 и более поздней версии для Windows
- Word 2016 и более поздней версии для Mac
- Word в Интернете
- Word для iPad

## <a name="imagecoercion-12"></a>ImageCoercion 1.2

ImageCoercion 1.2 позволяет преобразования в формат SVG () при записи данных `Office.CoercionType.XmlSvg` с помощью [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) метода. Поддерживаются следующие приложения.

- Excel на Windows (подключен к подписке Microsoft 365)
- Excel Mac (подключен к подписке Microsoft 365)
- PowerPoint на Windows (подключен к подписке Microsoft 365)
- PowerPoint Mac (подключен к подписке Microsoft 365)
- PowerPoint в Интернете
- Word on Windows (подключен к подписке Microsoft 365)
- Word на Mac (подключен к подписке Microsoft 365)

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md)
- [Указание приложений Office и обязательных элементов API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)
