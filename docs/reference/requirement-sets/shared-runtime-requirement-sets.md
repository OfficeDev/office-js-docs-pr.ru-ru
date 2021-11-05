---
title: Общие наборы требований к времени запуска
description: Указывает платформы и Office приложения, поддерживающую API SharedRuntime.
ms.date: 11/03/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: a5f7d3c9394de047b358d7f190c5adae5b5199b1
ms.sourcegitcommit: 210251da940964b9eb28f1071977ea1fe80271b4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/05/2021
ms.locfileid: "60793605"
---
# <a name="shared-runtime-requirement-sets"></a>Общие наборы требований к времени запуска

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

Части надстройки Office JavaScript, например области задач, файлы функций, запущенные из команд надстройки, и Excel настраиваемые функции, могут совместно использовать одно время запуска JavaScript. Это позволяет всем частям обмениваться набором глобальных переменных, обмениваться набором загруженных библиотек и общаться друг с другом без необходимости передавать сообщения через сохраняемую хранилище. Дополнительные сведения см. в Office надстройки для использования общего времени [работы JavaScript.](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)

В следующей таблице перечислены набор требований SharedRuntime 1.1, Office клиентские приложения, поддерживают этот набор требований, а также номера сборки или версии для Office приложения.

| Набор обязательных элементов | Office 2021 или более поздней Windows<br>(единовременная покупка) | Office для Windows<br>(подключено к подписке на Microsoft 365) | Office для iPad<br>(подключено к подписке на Microsoft 365) | Office для Mac<br>(подключено к подписке на Microsoft 365) | Office в Интернете | Office Online Server |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| SharedRuntime 1.1  | Сборка 16.0.14326.20454 или более поздней | Версия 2002 (сборка 12527.20092) или более поздней версии | Н/Д | 16.35 или более поздняя | Февраль 2020 г. | Н/Д |

> [!IMPORTANT]
> В настоящее время общая среда выполнения JavaScript не поддерживается на iPad или в версиях Office 2019 (или более ранних), предлагаемых в виде единовременных покупок. Дополнительные сведения о поддержке см. в следующих разделах.

## <a name="support-for-version-11-on-excel"></a>Поддержка версии 1.1 на Excel

Набор требований SharedRuntime 1.1 выпущен для Excel в Интернете, Windows и Mac.

## <a name="preview-support-for-version-11-on-word-and-powerpoint"></a>Поддержка предварительного просмотра версии 1.1 в Word и PowerPoint

В следующей таблице перечислены дополнительные сборки приложений, которые поддерживают предварительный просмотр общего времени запуска JavaScript. Версия предварительного просмотра общего времени работы подлежит изменению. Ее применение не поддерживается в рабочих средах. Чтобы получить новейшую сборку, вам нужно [присоединиться к программе предварительной оценки Office](https://insider.office.com/join). Хороший способ ознакомиться с такими возможностями — использование подписки на Microsoft 365. Если у вас еще нет подписки на Microsoft 365, вы можете оформить ее, присоединившись к [программе для разработчиков Microsoft 365](https://developer.microsoft.com/office/dev-program).

|Приложение Office |Сборка |
|-------------------|------|
|PowerPoint для Windows |Сборка 16.0.13218.10000 или более поздней |
|Word для Windows |Сборка 16.0.13218.10000 или более поздней |
|Word для Mac |Сборка 16.46.207.0 или более поздней |

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
