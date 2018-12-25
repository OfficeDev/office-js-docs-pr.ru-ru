---
title: Наборы обязательных элементов API диалоговых окон
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: ad0d472ebdcbdb9d61e78f6bdc9bfe7c08311cd7
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432657"
---
# <a name="dialog-api-requirement-sets"></a>Наборы обязательных элементов API диалоговых окон

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Надстройки Office работают в нескольких версиях Office. В приведенной ниже таблице перечислены наборы обязательных элементов Dialog API, ведущие приложения Office, которые их поддерживают, а также номера сборок или версий для этих приложений.

|  Набор обязательных элементов  | Office 2013 для Windows | Office 2016 для Windows (установка с помощью MSI)   | Office 365 для Windows (установка с помощью технологии "нажми и работай")   |  Office 365 для iPad  |  Office 365 для Mac  | Office Online  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.1  | Сборка 15.0.4855.1000 или более поздняя | Сборка 16.0.4390.1000 или более поздняя | Версия 1602 (сборка 6741.0000) или более поздняя | 1.22 или более поздняя | 15.20 или более поздняя| Январь 2017 г. | Версия 1608 (сборка 7601.6800) или более поздняя|

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

- [Номера версий и сборок выпусков из канала обновления для клиентов Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);
- [Какая у меня версия Office](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19);
- [Где можно найти номера версии и сборки клиентского приложения Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);
- [Обзор Office Online Server](https://docs.microsoft.com/officeonlineserver/office-online-server-overview).

## <a name="office-common-api-requirement-sets"></a>Стандартные наборы обязательных элементов API для Office

Сведения о типичных наборах обязательных элементов API см. в статье [Стандартные наборы обязательных элементов API для Office](office-add-in-requirement-sets.md).

## <a name="dialog-api-11"></a>Dialog API 1.1 

Dialog API 1.1 — это первая версия этого API. Дополнительные сведения об этом API см. в справочной статье о [Dialog API](/javascript/api/office/office.ui).

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Указание ведущих приложений Office и обязательных элементов API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-манифест надстроек Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
