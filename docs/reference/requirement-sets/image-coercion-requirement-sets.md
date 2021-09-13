---
title: Наборы обязательных элементов для приведения изображений
description: Поддержка наборов требований к принуждению изображений с Office надстройки в Excel, PowerPoint и Word.
ms.date: 02/19/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: 1e55eba4d28b459f4ffe9d402640dd04cff9acb4
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150626"
---
# <a name="image-coercion-requirement-sets"></a>Наборы обязательных элементов для приведения изображений

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

## <a name="imagecoercion-11"></a>ImageCoercion 1.1

ImageCoercion 1.1 позволяет преобразования в изображение () при записи `Office.CoercionType.Image` данных с помощью [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__options__callback_) метода. Поддерживаются следующие приложения.

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

ImageCoercion 1.2 позволяет преобразования в формат SVG () при записи данных `Office.CoercionType.XmlSvg` с помощью [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__options__callback_) метода. Поддерживаются следующие приложения.

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
