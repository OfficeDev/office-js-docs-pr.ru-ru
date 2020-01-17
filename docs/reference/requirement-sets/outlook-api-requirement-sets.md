---
title: Наборы обязательных элементов API JavaScript для Outlook
description: ''
ms.date: 01/14/2020
ms.prod: outlook
localization_priority: Priority
ms.openlocfilehash: bd6b20e9f0ddb5141f2f889a4e99af2c042a10ab
ms.sourcegitcommit: 212c810f3480a750df779777c570159a7f76054a
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/17/2020
ms.locfileid: "41217374"
---
# <a name="outlook-javascript-api-requirement-sets"></a>Наборы обязательных элементов API JavaScript для Outlook

Для надстроек Outlook требуются определенные версии API, которые указываются в элементе Requirements в их манифесте. Надстройки Outlook всегда включают элемент Set с атрибутом , для которого задано значение , и атрибутом , для которого установлен минимальный набор требований API, поддерживающий сценарии надстройки.

Например, в следующем фрагменте манифеста указан минимальный набор обязательных элементов 1.1.

```xml
<Requirements>
  <Sets>
    <Set Name="Mailbox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

Все API-интерфейсы Outlook приведены в `Mailbox`[наборе требований](../../develop/specify-office-hosts-and-api-requirements.md). У набора требований `Mailbox` есть версии, а каждый новый выпускаемый набор API-интерфейсов приведен в наборе более поздней версии. Не все клиенты Outlook поддерживают новейший набор API-интерфейсов, но если для клиента Outlook объявлена поддержка набора требований, то он будет поддерживать все API-интерфейсы в этом наборе.

Задайте версию минимального набора требований в манифесте, чтобы указать клиент Outlook, в котором появится надстройка. Если клиент не поддерживает минимальный набор требований, он не загружает надстройку. Например, если указана версия набора требований 1.3, надстройка не отобразится в каком-либо клиенте Outlook, который не поддерживает версии 1.3. и ниже

> [!NOTE]
> Чтобы использовать API в любом из нумерованных наборов обязательных элементов, следует ссылаться на **рабочую** библиотеку в сети CDN (https://appsforoffice.microsoft.com/lib/1/hosted/office.js).
>
> Сведения об использовании API предварительных версий см. в разделе [Использование предварительных версий API](#using-preview-apis) далее в этой статье.

## <a name="using-apis-from-later-requirement-sets"></a>Использование API из наборов обязательных элементов более поздних версий

Установка набора обязательных элементов не ограничивает доступные API, которые может использовать надстройка. Например, если для надстройки указан набор обязательных элементов "Mailbox 1.1", но она выполняется в клиенте Outlook, который поддерживает набор "Mailbox 1.3", надстройка может использовать API из набора обязательных элементов "Mailbox 1.3".

Чтобы использовать более новые API, разработчики могут проверить, поддерживает ли ведущее приложение набор обязательных элементов, выполнив следующее.

```js
if (Office.context.requirements.isSetSupported('Mailbox', '1.3')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

Кроме того, разработчики могут проверить наличие более новых API с помощью стандартных методов JavaScript.

```js
if (item.somePropertyOrFunction !== undefined) {
  // Use item.somePropertyOrFunction.
  item.somePropertyOrFunction;
}
```

Такие проверки не нужно выполнять для API-интерфейсов, присутствующих в версии набора обязательных элементов, указанной в манифесте.

## <a name="choosing-a-minimum-requirement-set"></a>Выбор минимального набора обязательных элементов

Разработчикам следует использовать набор обязательных элементов самой ранней версии, содержащий набор критически важных API для сценария их работы, без которого надстройка не будет работать.

## <a name="requirement-sets-supported-by-exchange-servers-and-outlook-clients"></a>Наборы обязательных элементов, поддерживаемые серверами Exchange и клиентами Outlook

В этом разделе указан диапазон наборов обязательных элементов, поддерживаемых сервером Exchange и клиентами Outlook.

> [!IMPORTANT]
> Если целевой сервер Exchange и клиент Outlook поддерживают разные наборы обязательных элементов, вы ограничены применением более низкой версии набора обязательных элементов. Например, если надстройка работает в Outlook 2016 для Mac (максимальная версия набора обязательных элементов: 1.6) с использованием Exchange 2013 (максимальная версия набора обязательных элементов: 1.1), ваша надстройка ограничивается применением набора обязательных элементов 1.1.

### <a name="exchange-server-support"></a>Поддержка сервера Exchange

Указанные ниже серверы поддерживают надстройки Outlook.

| Продукт | Основная версия Exchange | Поддерживаемые наборы обязательных элементов API |
|---|---|---|
| Exchange Online | Последняя сборка | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md) |
| Локальная среда Exchange | 2019 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) |
|| 2016 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) |
|| 2013 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="outlook-client-support"></a>Поддержка клиентов Outlook

Надстройки поддерживаются в Outlook на следующих платформах.

| Платформа | Основная версия Office или Outlook | Подписка или единовременная покупка? | Поддерживаемые наборы обязательных элементов API |
|---|---|---|---|
| Windows | Последние сборки<br>(Monthly channel) | Подписка на Office 365 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md) |
|| 2019 | Единовременная покупка | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md) |
|| 2016 | Единовременная покупка | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md) |
|| 2013 | Единовременная покупка | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md) |
| Mac | Последние сборки<br>(Monthly channel) | Подписка на Office 365 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md) |
|| 2019 | Единовременная покупка | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md) |
|| 2016 | Единовременная покупка | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md) |
| iOS | Последние сборки<br>(Monthly channel) | Подписка на Office 365 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) |
| Android | Последние сборки<br>(Monthly channel) | Подписка на Office 365 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) |
| Веб-браузер | современная версия | Exchange Online: подписка на Office 365, Outlook.com | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md) |
|| классическая версия | Локальная среда Exchange | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md) |

> [!NOTE]
> Поддержка версии 1.3 в Outlook 2013 добавлена в рамках [обновления для Outlook 2013 (KB3114349) от 8 декабря 2015 г.](https://support.microsoft.com/kb/3114349) Поддержка версии 1.4 в Outlook 2013 добавлена в рамках [обновления для Outlook 2013 (KB3118280) от 13 сентября 2016 г.](https://support.microsoft.com/help/3118280) Поддержка версии 1.4 в Outlook 2016 (MSI) добавлена в рамках [обновления для Office 2016 (KB4022223) от 3 июля 2018 г.](https://support.microsoft.com/help/4022223).

> [!TIP]
> Классическую и современную версии Outlook в веб-браузере можно различить по внешнему виду панели инструментов почтового ящика.
>
> **современная версия**
>
> ![снимок части экрана с изображением панели инструментов современной версии Outlook](https://docs.microsoft.com/outlook/add-ins/images/outlook-on-the-web-new-toolbar.png)
>
> **классическая версия**
>
> ![снимок части экрана с изображением панели инструментов классической версии Outlook](https://docs.microsoft.com/outlook/add-ins/images/outlook-on-the-web-classic-toolbar.png)

## <a name="using-preview-apis"></a>Использование предварительных версий API

Новые API JavaScript для Outlook сначала выпускаются в "предварительной версии", а затем становятся частью определенного нумерованного набора обязательных элементов после выполнения достаточного тестирования и получения отзывов пользователей. Чтобы отправить отзыв о предварительной версии API, используйте способ обратной связи, представленный в конце веб-страницы с описанием API.

> [!NOTE]
> API предварительной версии могут быть изменены и не предназначены для использования в рабочей среде.

Дополнительные сведения о предварительных версиях интерфейсов API см. в статье [Предварительная версия набора обязательных элементов API для Outlook](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md).
