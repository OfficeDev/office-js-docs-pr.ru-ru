---
title: Наборы обязательных элементов для приведения изображений
description: Поддержка наборов требований для приведения изображений с надстройками Office в Excel, PowerPoint и Word.
ms.date: 08/13/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 59f6891182f47bed1b7e3b6aa69a30e941bce7cb
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094357"
---
# <a name="image-coercion-requirement-sets"></a>Наборы обязательных элементов для приведения изображений

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

## <a name="imagecoercion-11"></a>ImageCoercion 1.1

Использовать imagecoercion 1,1 обеспечивает преобразование в Image ( `Office.CoercionType.Image` ) при записи данных с помощью [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) метода. Поддерживаются следующие узлы:

- Excel 2013 и более поздних версий в Windows
- Excel 2016 и более поздних версий на компьютерах Mac
- Excel на iPad
- OneNote в Интернете
- PowerPoint 2013 и более поздних версий в Windows
- PowerPoint 2016 и более поздних версий на компьютерах Mac
- PowerPoint в Интернете
- PowerPoint на iPad
- Word 2013 и более поздней версии для Windows
- Word 2016 и более поздней версии для Mac
- Word в Интернете
- Word для iPad

## <a name="imagecoercion-12"></a>ImageCoercion 1.2

Использовать imagecoercion 1,2 обеспечивает преобразование в формат SVG ( `Office.CoercionType.XmlSvg` ) при записи данных с помощью [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) метода. Поддерживаются следующие узлы:

- Excel в Windows (подключено к подписке Microsoft 365)
- Excel на Mac (подключено к подписке Microsoft 365)
- PowerPoint в Windows (подключено к подписке Microsoft 365)
- PowerPoint на Mac (с подключением к подписке Microsoft 365)
- PowerPoint в Интернете
- Word в Windows (подключены к подписке Microsoft 365)
- Word на Mac (подключено к подписке Microsoft 365)
- Word в Интернете

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md)
- [Указание ведущих приложений Office и обязательных элементов API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)
