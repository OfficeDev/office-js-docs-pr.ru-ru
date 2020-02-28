---
title: Общие наборы требований среды выполнения
description: Указывает платформы и узлы Office, которые поддерживают API Шаредрунтиме.
ms.date: 02/11/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: dbb9d908154da074eaff6901c778adea168504a9
ms.sourcegitcommit: 7464eac3b54a6a6b65e27549a3ad603af6ee1011
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42315884"
---
# <a name="shared-runtime-requirement-sets"></a>Общие наборы требований среды выполнения

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Части надстройки Office, в которых выполняется код JavaScript, например области задач, файлы функций, запущенные из команд надстроек, и пользовательские функции Excel, могут совместно использовать одну среду выполнения JavaScript. Это позволяет всем частям совместно использовать набор глобальных переменных, для совместного использования набора загруженных библиотек и для общения друг с другом без необходимости передачи сообщений через постоянное хранилище.

В следующей таблице перечислены наборы требований Шаредрунтиме 1,1, ведущие приложения Office, которые поддерживают этот набор требований, а также номера сборок или версий приложений Office.

|  Набор обязательных элементов  |  Office 2013 (или более поздней версии) в Windows<br>(единовременная покупка) | Office для Windows<br>(версия, подключенная к подписке на Office 365)   |  Office для iPad<br>(версия, подключенная к подписке на Office 365)  |  Office для Mac<br>(версия, подключенная к подписке на Office 365)  | Office в Интернете  | Office Online Server |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Шаредрунтиме 1,1  | Н/Д | Версия 2002 (сборка 12527,20092) или более поздняя | Н/Д | 16,35 или более поздняя версия | Февраль 2020 г. | Н/Д |

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

- [Номера версий и сборок выпусков из канала обновления для клиентов Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);
- [Какая у меня версия Office](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19);
- [Где можно найти номера версии и сборки клиентского приложения Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);
- [Обзор Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Указание ведущих приложений Office и обязательных элементов API](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-манифест надстроек Office](/office/dev/add-ins/develop/add-in-manifests)
