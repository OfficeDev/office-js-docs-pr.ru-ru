---
title: Общие наборы требований среды выполнения
description: Задает платформы и приложения Office, которые поддерживают API Шаредрунтиме.
ms.date: 07/10/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 872277488dd8d26241d9b445200f429aa102e26e
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293466"
---
# <a name="shared-runtime-requirement-sets"></a>Общие наборы требований среды выполнения

Наборы требований — это именованные группы элементов API. Надстройки Office используют наборы требований, указанные в манифесте, или используют проверку среды выполнения, чтобы определить, поддерживает ли приложение Office API, необходимые надстройке. Более подробную информацию можно узнать в статье [версии Office и наборах требований](../../develop/office-versions-and-requirement-sets.md).

Части надстройки Office, в которых выполняется код JavaScript, например области задач, файлы функций, запущенные из команд надстроек, и пользовательские функции Excel, могут совместно использовать одну среду выполнения JavaScript. Это позволяет всем частям совместно использовать набор глобальных переменных, для совместного использования набора загруженных библиотек и для общения друг с другом без необходимости передачи сообщений через постоянное хранилище.

В следующей таблице перечислены наборы требований Шаредрунтиме 1,1, клиентские приложения Office, которые поддерживают этот набор требований, а также номера сборок или версий для приложения Office.

|  Набор обязательных элементов  |  Office 2013 (или более поздней версии) в Windows<br>(единовременная покупка) | Office для Windows<br>(подключено к подписке на Microsoft 365)   |  Office для iPad<br>(подключено к подписке на Microsoft 365)  |  Office для Mac<br>(подключено к подписке на Microsoft 365)  | Office в Интернете  | Office Online Server |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Шаредрунтиме 1,1  | Недоступно | Версия 2002 (сборка 12527,20092) или более поздняя | Недоступно | 16.35 или более поздняя | Февраль 2020 г. | Недоступно |

## <a name="office-versions-and-build-numbers"></a>Номера версий и сборок Office

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Обзор Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md)
- [Указание приложений Office и требований к API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)
