---
title: Наборы обязательных элементов API JavaScript для PowerPoint
description: Узнайте больше о наборах обязательных элементов API JavaScript для PowerPoint
ms.date: 07/10/2020
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: eebcc78e69cd35732853daaee32f36df2b37252e
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159264"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a>Наборы обязательных элементов API JavaScript для PowerPoint

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

В приведенной ниже таблице перечислены наборы обязательных элементов для PowerPoint, ведущие приложения Office, которые их поддерживают, а также версии сборок или даты выхода.

|  Набор обязательных элементов  |  Office для Windows<br>(подключено к подписке на Microsoft 365)  |  Office для iPad<br>(подключено к подписке на Microsoft 365)  |  Office для Mac<br>(подключено к подписке на Microsoft 365)  | Office в Интернете |
|:-----|-----|:-----|:-----|:-----|:-----|
| PowerPointApi 1.1 | Версия 1810 (сборка 11001.20074) или более поздняя | 2.17 или более поздняя | 16.19 или более поздняя | Октябрь 2018 г. |

## <a name="office-versions-and-build-numbers"></a>Номера версий и сборок Office

Дополнительные сведения о номерах версий и сборок Office см. в следующих статьях:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="powerpoint-javascript-api-11"></a>API JavaScript для PowerPoint 1.1

API JavaScript для PowerPoint 1.1 включает один API для создания новой презентации. Дополнительные сведения об API см. в статье [API JavaScript для PowerPoint](../../powerpoint/powerpoint-add-ins.md).

## <a name="runtime-requirement-support-check"></a>Проверка поддержки обязательных элементов в среде выполнения

В среде выполнения надстройки могут проверять, поддерживает ли ведущее приложение набор обязательных элементов API, выполняя следующую проверку.

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a>Проверка поддержки обязательных элементов в манифесте

Используйте элемент `Requirements` в манифесте надстройки, чтобы указать ключевые наборы обязательных элементов или элементы API, которые должна использовать надстройка. Если платформа или ведущее приложение Office не поддерживает наборы обязательных элементов или элементы API, указанные в элементе `Requirements`, надстройка не будет работать в этом ведущем приложении или на этой платформе, а также не будет отображаться в разделе "Мои надстройки".

Ниже показана надстройка, которая загружается во всех ведущих приложениях Office, поддерживающих набор обязательных элементов OneNoteApi версии 1.1.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Большинство функций надстройки PowerPoint определяются набором обязательных элементов общего API. Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для PowerPoint](/javascript/api/powerpoint)
- [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md)
- [Указание ведущих приложений Office и обязательных элементов API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)
