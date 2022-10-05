---
title: Контекстные надстройки Outlook
description: Инициируют задачи, связанные с сообщением, не выходя из самого сообщения, что обеспечивает большее удобство и богатый пользовательский опыт
ms.date: 10/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: 73a13787dac7a6e74db6b919cc01a6dd33d29ab5
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467030"
---
# <a name="contextual-outlook-add-ins"></a>Контекстные надстройки Outlook

Contextual add-ins are Outlook add-ins that activate based on text in a message or appointment. By using contextual add-ins, a user can initiate tasks related to a message without leaving the message itself, which results in an easier and richer user experience.

[!include[JSON manifest does not support contextual add-ins](../includes/json-manifest-outlook-contextual-not-supported.md)]

Ниже приведены примеры контекстных надстроек.

- Чтобы открыть карту расположения, нужно выбрать адрес.
- Чтобы открыть надстройку с предложенными собраниями, нужно щелкнуть соответствующую строку.
- Чтобы добавить номер телефона в контакты, нужно его выбрать.


> [!NOTE]
> Контекстные надстройки в настоящее время недоступны в Outlook для Android и iOS, но в будущем мы добавим эту функцию.
>
> Поддержка этой функции была реализована в наборе обязательных элементов 1.6 См [клиенты и платформы](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.

## <a name="how-to-make-a-contextual-add-in"></a>Создание контекстной надстройки

Манифест контекстной надстройки должен содержать элемент [ExtensionPoint](/javascript/api/manifest/extensionpoint#detectedentity), для атрибута `xsi:type` которого задано значение `DetectedEntity`. В элементе **\<ExtensionPoint\>** надстройка указывает сущности или регулярное выражение, которые могут ее активировать. Если указывается сущность, это может быть любое из свойств объекта [Entities](/javascript/api/outlook/office.entities).

Следовательно, манифест надстройки должен содержать правило типа **ItemHasKnownEntity** или **ItemHasRegularExpressionMatch**. В следующем примере показано, как указать, что надстройка должна активировать сообщения с обнаруженной сущностью, которая является номером телефона.

```XML
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="contextLabel" />
  <!--If you opt to include RequestedHeight, it must be between 140px to 450px, inclusive.-->
  <!--<RequestedHeight>360</RequestedHeight>-->
  <SourceLocation resid="detectedEntityURL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="PhoneNumber" Highlight="all" />
  </Rule>
</ExtensionPoint>
```

После того как контекстная надстройка будет связана с учетной записью, она будет автоматически запускаться, когда пользователь щелкает выделенный объект или регулярное выражение. Дополнительные сведения о регулярных выражениях для надстроек Outlook см. в статье [Использование правил активации на основе регулярных выражений для отображения надстройки Outlook](use-regular-expressions-to-show-an-outlook-add-in.md).

Контекстные надстройки имеют несколько ограничений:

- Контекстная надстройка может существовать только в надстройках чтения (не надстройках создания).
- Указывать цвет выделенной сущности невозможно.
- Невыделенная сущность не запускает контекстную надстройку на карточке.

Так как невыделенные сущности и регулярные выражения не запускают контекстную надстройку, такие надстройки должны включать по крайней мере один элемент `Rule`, для атрибута `Highlight` которого задано значение `all`.

> [!NOTE]
> The `EmailAddress` and `Url` entity types do not support highlighting, so they cannot be used to launch a contextual add-in. They can however be combined in a `RuleCollection` rule type as an additional activation criteria.

## <a name="how-to-launch-a-contextual-add-in"></a>Запуск контекстной надстройки

A user launches a contextual add-in through text, either a known entity or a developer's regular expression. Typically, a user identifies a contextual add-in because the entity is highlighted. The following example shows how highlighting appears in a message. Here the entity (an address) is colored blue and underlined with a dotted blue line. A user launches the contextual add-in by clicking the highlighted entity. 

**Пример текста с выделенной сущностью (адресом)**

![Отображает выделенную сущность в сообщении электронной почты.](../images/outlook-detected-entity-highlight.png)
    
Если сообщение содержит несколько сущностей или контекстных надстроек, действует несколько правил взаимодействия:

- Если имеется несколько сущностей, нужно щелкнуть другую сущность, чтобы запустить связанную с ней надстройку.
- Если сущность активирует несколько надстроек, каждая из них открывается в новой вкладке. Пользователь может переключаться между надстройками, переходя между вкладками. Например, имя и адрес могут запускать надстройку для звонков и карту.
- If a single string contains multiple entities that activate multiple add-ins, the entire string is highlighted, and clicking the string shows all add-ins relevant to the string on separate tabs. For example, a string that describes a proposed meeting at a restaurant might activate the Suggested Meeting add-in and a restaurant rating add-in.

## <a name="how-a-contextual-add-in-displays"></a>Отображение контекстной надстройки

An activated contextual add-in appears in a card, which is a separate window near the entity. The card will normally appear below the entity and centered with respect to the entity as much as possible. If there is not enough room below the entity, the card is placed above it. The following screenshot shows the highlighted entity, and below it, an activated add-in (Bing Maps) in a card.

**Пример надстройки, отображаемой на карточке**

![Показывает контекстное приложение на карте.](../images/outlook-detected-entity-card.png)

Чтобы закрыть карточку и завершить работу надстройки, щелкните в любом месте за пределами карточки.

## <a name="current-contextual-add-ins"></a>Текущие контекстные надстройки

Следующие контекстные надстройки устанавливаются по умолчанию для пользователей с надстройки Outlook.

- Карты Bing
- Рекомендуемые собрания

## <a name="see-also"></a>См. также

- [Надстройка Outlook: номер заказа Contoso](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) (пример контекстной надстройки, которая активируется на основе соответствия регулярному выражению)
- [Написание вашей первой надстройки Outlook](../quickstarts/outlook-quickstart.md)
- [Использование правил активации на основе регулярных выражений для отображения надстройки Outlook](use-regular-expressions-to-show-an-outlook-add-in.md)
- [Объект Entities](/javascript/api/outlook/office.entities)
