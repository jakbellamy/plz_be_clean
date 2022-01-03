import dataset
import pandas as pd


class Database:
    def __init__(self, db_url=None):
        self.db_url = db_url
        self.db = None if not db_url else dataset.connect(db_url)
        self.engine = None if not db_url else self.db.engine

    def fetch_table(self, table_name):
        if not self.db:
            return False
        table = self.db[table_name].find()
        return pd.DataFrame(list(table))

    def to_sql(self, df, table_name, if_exists='append'):
        try:
            df.to_sql(table_name,
                      con=self.engine,
                      if_exists=if_exists,
                      chunksize=1000,
                      method='multi',
                      index=False)
            return True
        except Exception as e:
            print(e)
            return False

    @staticmethod
    def fill_foreign_value(child, child_reference, parent, parent_value, child_value=None):
        child_value = child_value if child_value else child_reference

        return child[child_value].apply(
            lambda x: parent.set_index('id').loc[x][parent_value] if x >= 0 else None)

    @staticmethod
    def fill_foreign_id(child, child_reference, parent, parent_value, child_value=None):
        child_value = child_value if child_value else child_reference

        return child[child_value].apply(
            lambda x: parent.set_index(parent_value).loc[x]['id'] if x else None)
