---
title: Наборы обязательных элементов API JavaScript для PowerPoint
description: Узнайте больше о наборах обязательных элементов API JavaScript для PowerPoint
ms.date: 10/26/2020
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: cf9ab510e4b35a140c77ee958279cb85a2189fa2
ms.sourcegitcommit: a4e09546fd59579439025aca9cc58474b5ae7676
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/27/2020
ms.locfileid: "48774730"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a>Наборы обязательных элементов API JavaScript для PowerPoint

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

В таблице ниже перечислены наборы обязательных элементов для PowerPoint, клиентские приложения Office, которые их поддерживают, а также версии сборок или даты выхода.

|  Набор обязательных элементов  |  Office для Windows<br>(подключено к подписке на Microsoft 365)  |  Office для iPad<br>(подключено к подписке на Microsoft 365)  |  Office для Mac<br>(подключено к подписке на Microsoft 365)  | Office в Интернете |
|:-----|-----|:-----|:-----|:-----|:-----|
| [Предварительная версия](powerpoint-preview-apis.md)  | Используйте последнюю версию Office, чтобы воспользоваться предварительными версиями API (может потребоваться присоединиться к [программе предварительной оценки Office](https://insider.office.com)). |
| PowerPointApi 1.1 | Версия 1810 (сборка 11001.20074) или более поздняя | 2.17 или более поздняя | 16.19 или более поздняя | Октябрь 2018 г. |

## <a name="office-versions-and-build-numbers"></a>Номера версий и сборок Office

Дополнительные сведения о номерах версий и сборок Office см. в следующих статьях:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="powerpoint-javascript-api-11"></a>API JavaScript для PowerPoint 1.1

API JavaScript для PowerPoint 1.1 содержит [единый API для создания новых презентаций](/javascript/api/powerpoint#powerpoint-createpresentation-base64file-). Сведения об этом API см. в разделе [Создание презентации](../../powerpoint/powerpoint-add-ins.md#create-a-presentation).

## <a name="how-to-use-powerpoint-requirement-sets-at-runtime-and-in-the-manifest"></a>Использование наборов обязательных элементов PowerPoint в среде выполнения и в манифесте

> [!NOTE]
> В этом разделе предполагается, что вы знакомы с общими сведениями о наборах обязательных элементов, изложенными в статьях [Версии и наборы обязательных элементов Office](../../develop/office-versions-and-requirement-sets.md) и [Указание приложений и обязательных элементов API Office](../../develop/specify-office-hosts-and-api-requirements.md).

Наборы требований — это именованные группы элементов API. Надстройка Office может выполнить проверку в среде выполнения или использовать указанные в манифесте наборы обязательных элементов, чтобы определить, поддерживает ли приложение Office необходимые надстройке API.

### <a name="checking-for-requirement-set-support-at-runtime"></a>Проверка поддержки наборов обязательных элементов в среде выполнения

В следующем примере кода показано, как определить, поддерживает ли приложение Office, в котором запускается надстройка, указанный набор обязательных элементов API.

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
} else {
  // Provide alternate flow/logic.
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a>Определение поддержки наборов обязательных элементов в манифесте

С помощью [элемента Requirements](../manifest/requirements.md) в манифесте надстройки можно указать минимальные наборы обязательных элементов и/или методы API, необходимые надстройке для активации. Если приложение или платформа Office не поддерживает наборы обязательных элементов или методы API, указанные в элементе манифеста `Requirements`, надстройка не будет работать в этом приложении или на этой платформе и не будет отображать список надстроек, показанный в разделе **Мои надстройки** . Если вашей надстройке для полной функциональности необходим определенный набор обязательных элементов, но она может быть полезна пользователям даже на тех платформах, которые не поддерживают этот набор, мы рекомендуем проверить поддержку обязательных элементов в среде выполнения как описано выше, а не прописывать поддержку набора обязательных элементов в манифесте.

В следующем примере кода показан элемент `Requirements` в манифесте надстройки, где указано, что надстройка должна загружаться во всех клиентских приложениях Office, поддерживающих набор обязательных элементов PowerPointApi версии 1.1 или более поздней.

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
- [Указание приложений Office и обязательных элементов API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)
