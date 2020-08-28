---
title: Наборы обязательных элементов API ленты
description: Указывает, какие платформы и сборки Office поддерживают динамические API ленты.
ms.date: 08/26/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: f734931817111ce52f779946e1f983ecc9238d3a
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293494"
---
# <a name="ribbon-api-requirement-sets"></a>Наборы обязательных элементов API ленты

Наборы требований — это именованные группы элементов API. Надстройки Office используют наборы требований, указанные в манифесте, или используют проверку среды выполнения, чтобы определить, поддерживает ли приложение Office API, необходимые надстройке. Более подробную информацию можно узнать в статье [версии Office и наборах требований](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Набор API ленты поддерживает программное управление при включении и отключении пользовательских команд надстроек (то есть кнопок ленты и их элементов меню).

Надстройки Office работают в нескольких версиях Office. В следующей таблице перечислены наборы требований API ленты, клиентские приложения Office, которые поддерживают этот набор требований, а также номера сборок или версий для приложения Office.

|  Набор обязательных элементов  | Office 2013 для Windows<br>(единовременная покупка) | Office 2016 или более поздней версии в Windows<br>(единовременная покупка)   | Office для Windows\*<br>(подключено к подписке на Microsoft 365) |  Office для iPad<br>(подключено к подписке на Microsoft 365)  |  Office для Mac\*<br>(подключено к подписке на Microsoft 365)  | Office в Интернете\*  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Риббонапи 1,1  | Н/Д | Н/Д | Ознакомьтесь со статьей поддержка<br>раздел ниже | Недоступно | 16,38 | Скоро | Недоступно|

> **&#42;** API ленты поддерживается только в Excel и требует подписки Microsoft 365. 

## <a name="office-on-windows-subscription-support"></a>Поддержка Office в Windows (подписка)

Набор обязательных элементов поддерживается в канале потребителей версии 2006 (сборка, 13001,20498 или выше). Для Office в Windows эта функция также поддерживается в течение полугодового канала и на ежемесячные сборки корпоративного канала, доступные в июле 14th 2020 или более поздней версии. Ниже приведены минимальные поддерживаемые сборки для каждого канала.  

|Канал | Версия | Сборка|
|:-----|:-----|:-----|
|Актуальный канал | 2006 или выше | 20266,20266 или выше|
|Ежемесячный канал (корпоративный) | 2005 или выше | 12827,20538 или выше|
|Ежемесячный канал (корпоративный) | 2004 | 12730,20602 или выше|
|Полугодовой канал (корпоративный) | 2002 или выше | 12527,20880 или выше|

## <a name="more-information"></a>Дополнительная информация

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

- [Номера версий и сборок для выпусков канала обновления для клиентов Microsoft 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Какая у меня версия Office](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19);
- [Где можно найти версию и номер сборки для клиентского приложения Microsoft 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Обзор Office Online Server](/officeonlineserver/office-online-server-overview)

> [!NOTE]
> Набор требований **риббонапи 1,1** пока не поддерживается в манифесте, поэтому его нельзя указать в разделе манифеста `<Requirements>` .


## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="ribbon-api-11"></a>API ленты 1,1

API ленты 1,1 — это первая версия API. Более подробную информацию об API можно узнать в разделе [Office. Ribbon ](/javascript/api/office/office.ribbon) .

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Указание приложений Office и требований к API](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-манифест надстроек Office](/office/dev/add-ins/develop/add-in-manifests)
