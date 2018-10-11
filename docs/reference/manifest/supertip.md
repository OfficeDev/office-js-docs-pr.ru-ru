# <a name="supertip"></a>Supertip

Определяет форматированную подсказку (элементы Title и Description). Используется элементами управления [Кнопка](control.md#button-control) или [Меню](control.md#menu-dropdown-button-controls).

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Description  |
|:-----|:-----|:-----|
|  [Title](#title)        | Да |   Текст суперподсказки.         |
|  [Description](#description)  | Да |  Описание суперподсказки.    |

### <a name="title"></a>Название

Обязательный элемент. Текст суперподсказки. Атрибуту **resid** должно быть присвоено значение атрибута **id** элемента **String** в элементе **ShortStrings**, вложенном в элемент [Resources](resources.md).

### <a name="description"></a>Description

Обязательный элемент. Описание суперподсказки. Атрибуту **resid** должно быть присвоено значение атрибута **id** элемента **String** в элементе **LongStrings**, вложенном в элемент [Resources](resources.md).

## <a name="example"></a>Пример

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
