This is a simple replacement for in-memory ADODB recordsets. It's reasonably fast and has no external dependencies.
Insert the class code into a new class module VBA project.

The sample code provides a range of examples.

The class supports:

  Fields: 
  
    Add (field names and field type), 
    Delete (including any column data), 
    Count (returns a count of fields), 
    Collection (returns a collection object containing the fields), 
    Exists (returns true if the field name exists),
    Index (returns the index of the named field)
    
  Records:
  
    Add (add a record using a variant array of field values),
    Retrieve (read a record to a variant array),
    Delete (delete using a row number or combine with the **Find** function), 
    DeleteAll (delete all of the records in the recordset),
    Count (returns a count of rows in the recordset), 
    Update (replaces a current record with new values).
    Collection (returns a collection object containing the records), 
    
  Field:
  
    Field(<row number>,<field name>) = <value> (sets a field value on a row),
    <value> = Field(<row number>,<field name>) (retrieves a field value on a row).
    
  Filter:

    Filter <field name>, <expression>, <value>** (filters based on an expression and value),
    
      Supported expressions include:
          Equals,
          GreaterThan,
          GreaterThanOrEqualTo,
          LessThan,
          LessThanOrEqualTo,
          UnequalTo,
          eLike (for wildcards),

      FilterOff (to restore the unfiltered records).

  Find:

      <row number> = Find(<field name>, <expression>, <search for value>, <FindFirst or FindNext>, <Forward or Backward>) (returns the row number of found, of 0 if no match)

  Sort:

      Sort <Field name>, <descending> (sort on any field)

  Debug Print:

      DebugPrint (writes the current recordset to the immediate window)
