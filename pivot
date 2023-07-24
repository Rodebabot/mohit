rows_field = input("Enter the name of the field for rows: ")
columns_field = input("Enter the name of the field for columns: ")
values_field = input("Enter the name of the field for values: ")
agg_func = input("Enter the aggregation function (e.g., sum, mean, max, min, count): ")

# Create a pivot table
pivot_table = df.pivot_table(values=values_field, index=rows_field, columns=columns_field, aggfunc=agg_func)

# Print the pivot table
print("\nPivot Table:")
print(pivot_table)
