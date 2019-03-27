---
title: Наборы обязательных элементов API JavaScript для Outlook
description: ''
ms.date: 03/19/2019
ms.prod: outlook
localization_priority: Priority
ms.openlocfilehash: a17d73d0c7bd7307440229504329cc231c3acf39
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870697"
---
# <a name="outlook-javascript-api-requirement-sets"></a>Наборы обязательных элементов API JavaScript для Outlook

Для надстроек Outlook требуются определенные версии API, которые указываются в элементе Requirements в их манифесте. Надстройки Outlook всегда включают элемент Set с атрибутом , для которого задано значение , и атрибутом , для которого установлен минимальный набор требований API, поддерживающий сценарии надстройки.

Например, в следующем фрагменте манифеста указан минимальный набор обязательных элементов 1.1:

```xml
<Requirements>
  <Sets>
    <Set Name="MailBox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

Все API-интерфейсы Outlook приведены в `Mailbox`[наборе требований](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements). У набора требований `Mailbox` есть версии, а каждый новый выпускаемый набор API-интерфейсов приведен в наборе более поздней версии. Не все клиенты Outlook поддерживают новейший набор API-интерфейсов, но если для клиента Outlook объявлена поддержка набора требований, то он будет поддерживать все API-интерфейсы в этом наборе.

Задайте версию минимального набора требований в манифесте, чтобы указать клиент Outlook, в котором появится надстройка. Если клиент не поддерживает минимальный набор требований, он не загружает надстройку. Например, если указана версия набора требований 1.3, надстройка не отобразится в каком-либо клиенте Outlook, который не поддерживает версии 1.3. и ниже

## <a name="using-apis-from-later-requirement-sets"></a>Использование API из наборов обязательных элементов более поздних версий

Установка набора обязательных элементов не ограничивает доступные API, которые может использовать надстройка. Например, если для надстройки указан набор обязательных элементов 1.1, но она выполняется в клиенте Outlook, который поддерживает версию 1.3, надстройка может использовать API из набора обязательных элементов 1.3.

Чтобы использовать более новые API, разработчики могут просто проверить их наличие с помощью стандартных методов JavaScript.

```js
if (item.somePropertyOrFunction !== undefined) {
  item.somePropertyOrFunction ...
}
```

Такие проверки не нужно выполнять для API-интерфейсов, присутствующих в версии набора обязательных элементов, указанной в манифесте.

## <a name="choosing-a-minimum-requirement-set"></a>Выбор минимального набора обязательных элементов

Разработчикам следует использовать набор обязательных элементов самой ранней версии, содержащий набор критически важных API для сценария их работы, без которого надстройка не будет работать.

## <a name="clients"></a>Клиенты

Указанные ниже клиенты поддерживают надстройки Outlook.

| Клиент | Поддерживаемые наборы обязательных элементов API |
| --- | --- |
| Outlook для Windows (Office 365) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6), [1.7](/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Outlook 2019 для Windows | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6), [1.7](/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Outlook 2016 для Windows | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4) |
| Outlook 2013 для Windows | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4) |
| Outlook для Mac (Office 365) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6) |
| Outlook 2019 для Mac | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6) |
| Outlook 2016 для Mac | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6) |
| Outlook для iOS | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5) |
| Outlook для Android | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5) |
| Outlook в Интернете (Office 365 и Outlook.com) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6), [1.7](/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Outlook в Интернете (классическая версия) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6) |
| Любой клиент Outlook, подключенный к локальному серверу Exchange 2019 | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5) |
| Любой клиент Outlook, подключенный к локальному серверу Exchange 2016 | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3) |
| Любой клиент Outlook, подключенный к локальному серверу Exchange 2013 | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1) |

> [!NOTE]
> Поддержка версии 1.3 в Outlook 2013 добавлена в рамках [обновления для Outlook 2013 (KB3114349) от 8 декабря 2015 г.](https://support.microsoft.com/kb/3114349) Поддержка версии 1.4 в Outlook 2013 добавлена в рамках [обновления для Outlook 2013 (KB3118280) от 13 сентября 2016 г.](https://support.microsoft.com/help/3118280) Поддержка версии 1.4 в Outlook 2016 (MSI) добавлена в рамках [обновления для Office 2016 (KB4022223) от 3 июля 2018 г.](https://support.microsoft.com/help/4022223).

## <a name="using-preview-apis"></a>Использование предварительных версий API

Новые API JavaScript для Outlook сначала выпускаются в "предварительной версии", а затем становятся частью определенного нумерованного набора обязательных элементов после выполнения достаточного тестирования и получения отзывов пользователей. Чтобы отправить отзыв о предварительной версии API, используйте способ обратной связи, представленный в конце веб-страницы с описанием API.

> [!NOTE]
> API предварительной версии могут быть изменены и не предназначены для использования в рабочей среде.

Дополнительные сведения о предварительных версиях интерфейсов API см. в статье [Предварительная версия набора обязательных элементов API для Outlook](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md).