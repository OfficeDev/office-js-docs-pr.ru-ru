---
title: Наборы обязательных элементов для приведения изображений
description: Поддержка наборов требований к принуждению к изображениям с Office надстройки в Excel, PowerPoint и Word.
ms.date: 09/08/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: 50533d179180eeef81825a97c9c39fda95af554f
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746918"
---
# <a name="image-coercion-requirement-sets"></a>Наборы обязательных элементов для приведения изображений

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

## <a name="imagecoercion-11"></a>ImageCoercion 1.1

ImageCoercion 1.1 позволяет преобразования в изображение (`Office.CoercionType.Image`) при записи данных с помощью [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) метода. Поддерживаются следующие приложения.

- Excel 2013 и более поздней Windows
- Excel 2016 и более поздней основе на Mac
- Excel на iPad
- OneNote в Интернете
- PowerPoint 2013 и более поздней Windows
- PowerPoint 2016 и позднее на Mac
- PowerPoint в Интернете
- PowerPoint на iPad
- Word 2013 и более поздней версии для Windows
- Word 2016 и более поздней версии для Mac
- Word в Интернете
- Word для iPad

## <a name="imagecoercion-12"></a>ImageCoercion 1.2

ImageCoercion 1.2 позволяет преобразования в формат SVG (`Office.CoercionType.XmlSvg`) при записи данных с помощью [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) метода. Поддерживаются следующие приложения.

- Excel 2021 г. и более поздней Windows
- Excel 2021 г. и позднее на Mac
- PowerPoint 2021 г. и более поздней Windows
- PowerPoint 2021 г. и более поздней
- PowerPoint в Интернете
- Word 2021 и более поздний Windows
- Word 2021 и более поздний на Mac

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md)
- [Указание приложений Office и обязательных элементов API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)
