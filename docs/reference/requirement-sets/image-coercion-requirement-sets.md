---
title: Наборы обязательных элементов для приведения изображений
description: Поддержка наборов требований для приведения изображений с надстройками Office в Excel, PowerPoint и Word.
ms.date: 08/13/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 7140099757c6e4b5ad405723d5fed95fded6d919
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293550"
---
# <a name="image-coercion-requirement-sets"></a>Наборы обязательных элементов для приведения изображений

Наборы требований — это именованные группы элементов API. Надстройки Office используют наборы требований, указанные в манифесте, или используют проверку среды выполнения, чтобы определить, поддерживает ли приложение Office API, необходимые надстройке. Более подробную информацию можно узнать в статье [версии Office и наборах требований](../../develop/office-versions-and-requirement-sets.md).

## <a name="imagecoercion-11"></a>ImageCoercion 1.1

Использовать imagecoercion 1,1 обеспечивает преобразование в Image ( `Office.CoercionType.Image` ) при записи данных с помощью [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) метода. Поддерживаются следующие приложения:

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

Использовать imagecoercion 1,2 обеспечивает преобразование в формат SVG ( `Office.CoercionType.XmlSvg` ) при записи данных с помощью [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) метода. Поддерживаются следующие приложения:

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
- [Указание приложений Office и требований к API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)
