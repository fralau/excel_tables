"Test Excel utilities"

import pytest
from excel_tables.excel_util import to_argb, get_font_color, expand_name

def test_colors():
    "Test all the color utilities"
    COLORS_TABLE = [
        ("Red", "FF0000"),
        ("darkolivegreen", "556b2f"),
        ("indigo", "4b0082"),
        ("Yellow", "FFFF00"),
        ("Cyan", "00FFFF"),
        ("Magenta", "FF00FF"),
        ("Orange", "FFA500"),
        ("Purple", "800080"),
        ("lightgoldenrodyellow", "fafad2"),
        ("powderblue", "b0e0e6")
    ]

    for name, argb in COLORS_TABLE:
        print(name, argb)
        assert to_argb(name).upper() == argb.upper()

    # test the color of the font against the background color
    # for title rows
    COLORS_BACKGROUND = [
        ('navy', 'white'),
        ('firebrick', 'white'),
        ('red', 'white'),
        ('black', 'white'),
        ('snow', 'black'),
        ('mintcream', 'black'),
        ('lightyellow', 'black')
    ]
    print("BACKGROUNDS:")
    for name, background in COLORS_BACKGROUND:
        print(name, background)
        assert get_font_color(name) == background

def test_expand_name():
    "Test expand name"

    actual = ['foo', 'bar', 'baz', 'foobar']
    assert expand_name('foo', actual) == ['foo']
    # regexp:
    assert expand_name('//foo', actual) == ['foo', 'foobar']
    assert expand_name('%foo', actual, regex_marker='%') == ['foo', 'foobar']

    # check errors
    with pytest.raises(ValueError) as excinfo:
        expand_name('bang', actual)
    assert "No item" in str(excinfo.value)
    assert 'bang' in str(excinfo.value)

    # give a description of the item:
    description='Thingamajig'
    with pytest.raises(ValueError) as excinfo:
        expand_name('bang', actual, desc=description)
    assert f"No {description}" in str(excinfo.value)


    assert expand_name('//foo', ['foo', 'bar', 'baz', 'foobar', 'No Fool']) \
        ==  ['foo', 'foobar', 'No Fool']

    # main use is to retrieve columns with similar names, e.g. date names
    columns = ['Name', 'Description', 'Record Date', 'Mydate', 
                'Length', 'Date of Birth']
    assert expand_name('//date', columns) == ['Record Date', 'Mydate',
                                                'Date of Birth']
    assert expand_name('//^date', columns) == ['Date of Birth']
    assert expand_name('//date', columns, 
                       case_sensitive=True) == ['Mydate']