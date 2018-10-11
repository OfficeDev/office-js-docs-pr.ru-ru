# <a name="metadata-element"></a>Элемент Metadata

Задает параметры метаданных, используемых настраиваемыми функциями в  Excel.

## <a name="attributes"></a>Атрибуты

Нет

## <a name="child-elements"></a>Дочерние элементы

|  Элемент  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Да  | Строка с идентификатором ресурса файла HTML, используемого настраиваемыми функциями. |

## <a name="example"></a>Пример

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
