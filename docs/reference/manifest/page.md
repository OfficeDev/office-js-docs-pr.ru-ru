# <a name="page-element"></a>Элемент Page

Определяет параметры HTML-страницы, используемые настраиваемой функцией в Excel.

## <a name="attributes"></a>Атрибуты

Нет

## <a name="child-elements"></a>Дочерние элементы

|  Элемент  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Да  | Строка с идентификатором ресурса HTML-файла, используемого настраиваемыми функциями. |

## <a name="example"></a>Пример

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
