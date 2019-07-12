---
title: Наборы требований к приведению изображений
description: Поддержка наборов требований для приведения изображений с надстройками Office в Excel, PowerPoint и Word.
ms.date: 07/11/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 046a3f1f16d8b48cddbd64bddf80a31ed1e50583
ms.sourcegitcommit: 61f8f02193ce05da957418d938f0d94cb12c468d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/11/2019
ms.locfileid: "35633993"
---
# <a name="image-coercion-requirement-sets"></a>Наборы требований к приведению изображений

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Надстройки Office работают в нескольких версиях Office. В приведенной ниже таблице перечислены наборы требований к приведению изображений, ведущие приложения Office, которые поддерживают этот набор требований, а также номера сборок или версий приложений Office.

## <a name="imagecoercion-11"></a>Использовать imagecoercion 1,1

Использовать imagecoercion 1,1 обеспечивает преобразование в Image (`Office.CoercionType.Image`) при записи данных с помощью [`Document.setSelectedDataAsync`](/javascript/api/office/document#setselecteddataasync-data--options--callback-) метода. Поддерживаются следующие узлы:

- Excel 2013 и более поздних версий в Windows
- Excel 2016 и более поздних версий на компьютерах Mac
- Excel в Интернете
- Excel на iPad
- OneNote в Интернете
- PowerPoint 2013 и более поздних версий в Windows
- PowerPoint 2016 и более поздних версий на компьютерах Mac
- PowerPoint в Интернете
- PowerPoint на iPad
- Word 2013 и более поздние версии для Windows
- Word 2016 и более поздние версии на компьютерах Mac
- Word в Интернете
- Word на iPad

## <a name="imagecoercion-12"></a>Использовать imagecoercion 1,2

Использовать imagecoercion 1,2 обеспечивает преобразование в формат SVG (`Office.CoercionType.XmlSvg`) при записи данных с помощью [`Document.setSelectedDataAsync`](/javascript/api/office/document#setselecteddataasync-data--options--callback-) метода. Поддерживаются следующие узлы:

- Excel в Windows (подключен к подписке на Office 365)
- Excel на Mac (подключен к подписке на Office 365)
- Excel в Интернете
- PowerPoint в Windows (подключено к подписке на Office 365)
- PowerPoint на Mac (с подключением к подписке на Office 365)
- PowerPoint в Интернете
- Word в Windows (подключен к подписке на Office 365)
- Word на Mac (подключен к подписке на Office 365)
- Word в Интернете

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Указание ведущих приложений Office и обязательных элементов API](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-манифест надстроек Office](/office/dev/add-ins/develop/add-in-manifests)
