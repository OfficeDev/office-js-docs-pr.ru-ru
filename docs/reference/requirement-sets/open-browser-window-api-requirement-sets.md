---
title: Открытие наборов требований окна браузера
description: Указывает, какие платформы и сборки Office поддерживают API Опенбровсервиндов.
ms.date: 09/16/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 8f6966f5bdcecd9c55a20f2d640d066906c1b6a3
ms.sourcegitcommit: 2479812e677d1a7337765fe8f1c8345061d4091a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/19/2020
ms.locfileid: "48135240"
---
# <a name="open-browser-window-api-requirement-sets"></a>Наборы обязательных элементов API окна браузера

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

Набор API Опенбровсервиндов позволяет надстройкам открывать браузер для выполнения задач, которые не всегда можно выполнить в изолированном элементе управления вебвиев внутри надстройки; Например, Загрузка PDF-файла, когда элемент управления вебвиев предоставляется Microsoft Edge.

Надстройки Office работают в нескольких версиях Office. В следующей таблице перечислены наборы требований API Опенбровсервиндов, ведущие приложения Office, которые поддерживают этот набор требований, а также номера сборок или версий для приложения Office.

|  Набор обязательных элементов  | Office 2013 в Windows или более поздней версии<br>(единовременная покупка) | Office для Windows<br>(версия, подключенная к подписке на Office 365) |  Office для iPad<br>(версия, подключенная к подписке на Office 365)  |  Office для Mac<br>(версия, подключенная к подписке на Office 365)  | Office в Интернете  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Опенбровсервиндовапи 1,1  | Недоступно | Версия 1810 (сборка 16.0.11001.20074) или более поздняя | 16.0.0.0 или более поздняя версия | 16.0.0.0 или более поздняя версия | Н/Д | Н/Д|

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

- [Номера версий и сборок выпусков из канала обновления для клиентов Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);
- [Какая у меня версия Office](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19);
- [Где можно найти номера версии и сборки клиентского приложения Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);
- [Обзор Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="openbrowserwindowapi-11"></a>Опенбровсервиндовапи 1,1

Опенбровсервиндовапи 1,1 — это первая версия API. Более подробную информацию об API можно узнать в разделе [Office. Context. UI](/javascript/api/office/office.context.ui) .

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md)
- [Указание ведущих приложений Office и обязательных элементов API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)
