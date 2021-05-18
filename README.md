# xls2pyobj

A module to convert rows of a spreadsheet like data into a list of python
objects based on a specification. Currently xls, xlsx and csv formats are
supported through a uniform interface. The file type is guessed by the file
extension only.

It is generally meant to be  useful for financial statements, but there is no
such restriction.

It is advisable to have a data model for one kind of statements. (Rather, when
various spread sheet formats feed to a common data model, this utility would
have a more meaningul use.) For example all bank statements can follow one data
model (field names etc., see contrib directory for examples) so that the
application can deal with statements of various banks seamlessly without having
to be customized for every source of data.

This is meant for simple standalone applications (say) for personal use and not
meant to scale for a large scale use. But if you have a large scale use, you
may use a database, and such utility to import spread sheets into your
database.

# Software requirements

- python3 is required

Following python packages are required:

- xlrd
- openpyxl
- csv

# Writing the json file describing the spread sheet format

See contrib directory to get an idea. Following is a description of the
properties involved:

    fields: Keys in the dictionary that follows become attribute names of the
    objects formed. Values capture the information required to construct the
    field values.

        col: Column number (1 based) in whchi to find the value of this field.

        trim: A list of characters to be removed from the data value. For
        example, "," is present in many financial spread sheets in amount
        fields and needs to be removed. Default is []

        typ: int|float|date|str, Default: str. Type of the field. The spread
        sheet data in a cell is converted into an object of this type.

    globals: Default properties of fields can be specified once here isntead of
    repeating with each field. For examle, datefmt. Besides, sheet level
    properties such as strtrow, endpat are to be specified here.

        datefmt: Format string for fields of date type as needed by datetime
        package.

        endpat: String that indicates end of the rows to be selected. Default
        is ''.

        endpatcol: Which column to look for to identify end of data. Default is
        1.

        strtrow: Row number (1 based) from which the actual data starts.

    remark: Just for your own documentation, ignored by the module



# contrib: Collection of spread sheet formats

The contrib directory contains specification of spread sheet formats of various
banks and other financial institutions. Of course one can add own json format.
If you think it may be useful to others, consider contributing.

# Usage

Following snippet may give some idea:

    from xls2obj import XlsObjs
    if __name__ == '__main__':
        objs = XlsObjs('mystmt.xls','format.json')
        for o in objs:
            # Do something
