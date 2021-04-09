---
title: Общие наборы требований к времени запуска
description: Указывает платформы и приложения Office, поддерживающую API SharedRuntime.
ms.date: 04/08/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 8d0db6e129aaf7a4aa2967e7a1341d6db1188359
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652225"
---
# <a name="shared-runtime-requirement-sets"></a>Общие наборы требований к времени запуска

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

Части надстройки Office, которая выполняет код JavaScript, такие как области задач, файлы функций, запущенные из команд надстройки, и настраиваемые функции Excel, могут совместно использовать одно время запуска JavaScript. Это позволяет всем частям обмениваться набором глобальных переменных, обмениваться набором загруженных библиотек и общаться друг с другом без необходимости передавать сообщения через сохраняемую хранилище. Дополнительные сведения см. в [раздел Настройка надстройки Office для использования](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)общего времени работы JavaScript.

В следующей таблице перечислены набор требований SharedRuntime 1.1, клиентские приложения Office, которые поддерживают этот набор требований, а также номера сборки или версии для приложения Office.

|  Набор обязательных элементов  |  Office 2013 (или более поздний) для Windows<br>(единовременная покупка) | Office для Windows<br>(подключено к подписке на Microsoft 365)   |  Office для iPad<br>(подключено к подписке на Microsoft 365)  |  Office для Mac<br>(подключено к подписке на Microsoft 365)  | Office в Интернете  | Office Online Server |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| SharedRuntime 1.1  | Недоступно | Версия 2002 (сборка 12527.20092) или более поздней версии | Недоступно | 16.35 или более поздняя | Февраль 2020 г. | Недоступно |

> [!IMPORTANT]
> Набор общих требований к времени работы JavaScript доступен только на следующих платформах.
>
> - Excel для Интернета, Windows и Mac.
> - PowerPoint для Windows (сборка 13218.10000 или более поздняя). Общая среда выполнения JavaScript для PowerPoint в настоящее время доступна в предварительной версии и может изменяться. Ее применение не поддерживается в рабочих средах. Чтобы получить новейшую сборку, вам нужно [присоединиться к программе предварительной оценки Office](https://insider.office.com/join). Хороший способ ознакомиться с такими возможностями — использование подписки на Microsoft 365. Если у вас еще нет подписки на Microsoft 365, вы можете оформить ее, присоединившись к [программе для разработчиков Microsoft 365](https://developer.microsoft.com/office/dev-program).
>
> В настоящее время общая среда выполнения JavaScript не поддерживается на iPad или в версиях Office 2019 (или более ранних), предлагаемых в виде единовременных покупок.

## <a name="office-versions-and-build-numbers"></a>Номера версий и сборок Office

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Обзор Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>См. также

- [Настройка надстройки Office для использования общей среды выполнения JavaScript](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md)
- [Указание приложений Office и обязательных элементов API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)
