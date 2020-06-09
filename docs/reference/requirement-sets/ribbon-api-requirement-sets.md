---
title: Наборы обязательных элементов API ленты
description: Указывает, какие платформы и сборки Office поддерживают динамические API ленты.
ms.date: 05/28/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 6625dbb0612b3ee79c5b97c0fd3aa28f214258a3
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611325"
---
# <a name="ribbon-api-requirement-sets"></a>Наборы обязательных элементов API ленты

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Набор API ленты поддерживает программное управление при включении и отключении пользовательских команд надстроек (то есть кнопок ленты и их элементов меню).

Надстройки Office работают в нескольких версиях Office. В следующей таблице перечислены наборы требований API ленты, ведущие приложения Office, которые поддерживают этот набор требований, а также номера сборок или версий приложения Office.

|  Набор обязательных элементов  | Office 2013 для Windows<br>(единовременная покупка) | Office 2016 или более поздней версии в Windows<br>(единовременная покупка)   | Office для Windows\*<br>(версия, подключенная к подписке на Office 365) |  Office для iPad<br>(версия, подключенная к подписке на Office 365)  |  Office для Mac\*<br>(версия, подключенная к подписке на Office 365)  | Office в Интернете\*  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Риббонапи 1,1  | Н/Д | Н/Д | Версия 2002 (сборка 12527,20264) или более поздняя | 16,38 или более поздняя версия | Н/Д | Февраль 2020 г. | Н/Д|

> **&#42;** На этапе предварительной версии API ленты поддерживается только в Excel и для него требуется Office 365 (версия подписки Office). Следует использовать последнюю версию для текущего месяца и сборку из канала для участников программы предварительной оценки. Чтобы получить эту версию, необходимо быть участником программы предварительной оценки Office. Дополнительные сведения см. на странице [Примите участие в программе предварительной оценки Office](https://products.office.com/office-insider?tab=tab-1). Обратите внимание, что при построении градуатес к производственному каналу, поддержка предварительных функций, в том числе API ленты, будет отключена для этой сборки.

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

- [Номера версий и сборок выпусков из канала обновления для клиентов Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);
- [Какая у меня версия Office](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19);
- [Где можно найти номера версии и сборки клиентского приложения Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);
- [Обзор Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="ribbon-api-11"></a>API ленты 1,1

API ленты 1,1 — это первая версия API. Более подробную информацию об API можно узнать в разделе [Office. Ribbon](/javascript/api/office/office.ribbon) .

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Указание ведущих приложений Office и обязательных элементов API](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-манифест надстроек Office](/office/dev/add-ins/develop/add-in-manifests)
