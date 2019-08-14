---
title: Наборы требований к приведению изображений
description: Поддержка наборов требований для приведения изображений с надстройками Office в Excel, PowerPoint и Word.
ms.date: 08/13/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 9d622c827315f6657cf0fddaace33968bd634d64
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395675"
---
# <a name="image-coercion-requirement-sets"></a>Наборы требований к приведению изображений

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

## <a name="imagecoercion-11"></a>Использовать imagecoercion 1,1

Использовать imagecoercion 1,1 обеспечивает преобразование в Image (`Office.CoercionType.Image`) при записи данных с помощью [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) метода. Поддерживаются следующие узлы:

- Excel 2013 и более поздних версий в Windows
- Excel 2016 и более поздних версий на компьютерах Mac
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

Использовать imagecoercion 1,2 обеспечивает преобразование в формат SVG (`Office.CoercionType.XmlSvg`) при записи данных с помощью [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) метода. Поддерживаются следующие узлы:

- Excel в Windows (подключен к подписке на Office 365)
- Excel на Mac (подключен к подписке на Office 365)
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
