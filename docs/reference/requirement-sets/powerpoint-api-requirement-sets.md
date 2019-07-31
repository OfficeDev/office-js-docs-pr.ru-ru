---
title: Наборы обязательных элементов API JavaScript для PowerPoint
description: ''
ms.date: 07/26/2019
ms.prod: powerpoint
localization_priority: Normal
ms.openlocfilehash: 4f64654a4130cc0d4bf96d9c59e364e77c808748
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/31/2019
ms.locfileid: "35941151"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a>Наборы обязательных элементов API JavaScript для PowerPoint

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

В следующей таблице перечислены наборы требований PowerPoint, ведущие приложения Office, которые поддерживают эти наборы требований, а также версии сборки или Дата доступности.

|  Набор обязательных элементов  |  Office для Windows<br>(подключено к подписке Office 365)  |  Office на iPad<br>(подключено к подписке Office 365)  |  Office на Mac<br>(подключено к подписке Office 365)  | Office в Интернете |
|:-----|-----|:-----|:-----|:-----|:-----|
| Поверпоинтапи 1,1 | Версия 1810 (сборка 11001,20074) или более поздняя | 2.17 или более поздняя | 16,19 или более поздняя версия | Октябрь 2018 г. |

## <a name="office-versions-and-build-numbers"></a>Номера версий и сборок Office

Более подробную информацию о версиях Office и номерах сборок можно узнать в следующих статьях:

- [Номера версий и сборок выпусков из канала обновления для клиентов Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);
- [Какая у меня версия Office](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19);
- [Где можно найти номера версии и сборки клиентского приложения Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);

## <a name="powerpoint-javascript-api-11"></a>API JavaScript для PowerPoint 1,1

API JavaScript для PowerPoint 1,1 содержит один API для создания новой презентации. Дополнительные сведения об API можно найти в статье [API JavaScript для PowerPoint](../../powerpoint/powerpoint-add-ins.md).

## <a name="runtime-requirement-support-check"></a>Проверка поддержки требований в среде выполнения

В среде выполнения надстройки могут проверять, поддерживает ли конкретный узел набор обязательных элементов API, выполнив следующие действия.

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a>Проверка поддержки обязательных элементов в манифесте

Используйте `Requirements` элемент в манифесте надстройки, чтобы указать критические наборы требований или элементы API, которые должна использовать надстройка. Если ведущее приложение или платформа Office не поддерживает наборы требований или элементы API, указанные в `Requirements` элементе, надстройка не будет запускаться на этом узле или платформе и не будет отображаться в папке "Мои надстройки".

Ниже показана надстройка, которая загружается во всех ведущих приложениях Office, поддерживающих набор обязательных элементов OneNoteApi версии 1.1.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Большинство функциональных возможностей надстройки PowerPoint берутся из общего набора API. Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для PowerPoint](/javascript/api/powerpoint)
- [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Указание ведущих приложений Office и обязательных элементов API](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-манифест надстроек Office](/office/dev/add-ins/develop/add-in-manifests)
