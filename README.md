# GUI-for-high-throughput-SQL-data-retriever
In some scenarios, we need to retrieve data from SQL database based on massive primary-key IDs. Usually, we can use the "WHERE ID IN ("ID1","ID2","ID3", ...)" clause to specify the records we want to retrieve. However, SQL can only handle limited amount of records at a time. When I built this app, I used a split-retrieve-concatenate strategy.

![image](https://github.com/hanfei1986/GUI-for-high-throughput-SQL-data-retriever/assets/59255164/7f26df9f-5bec-4766-ace5-286b98ace847)
