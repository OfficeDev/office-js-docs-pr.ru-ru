# <a name="customtab-element"></a>Элемент CustomTab

На ленте можно указать вкладку и группу для команд надстройки. Это может быть вкладка по умолчанию (**Главная**, **Сообщение** или **Собрание**) либо настраиваемая вкладка, которую определяет надстройка.

На настраиваемых вкладках надстройка может создать до 10 групп. Каждая группа может включать не более 6 элементов управления, независимо от того, на какой вкладке она отображается. Надстройка может создать не более одной настраиваемой вкладки.

Атрибут **id** должен быть уникальным для манифеста.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Group](group.md)      | Да |  Определяет группу команд.  |
|  [Label](#label-tab)      | Да |  Метка элемента CustomTab или Group.  |
|  [Control](control.md)    | Да |  Коллекция из одного или нескольких объектов Control.  |

### <a name="group"></a>Group

Обязательный. См. статью [элемент Group](group.md).

### <a name="label-tab"></a>Label (Tab)

Обязательный элемент. Метка настраиваемой вкладки. Атрибуту **resid** нужно присвоить значение атрибута **id** элемента **String** в элементе **ShortStrings**, входящем в элемент [Resources](resources.md).


## <a name="customtab-example"></a>Пример элемента CustomTab

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```