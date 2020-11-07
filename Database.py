import sqlite3 as sq


class DbConnection:

    """
    A class for connecting to sqlite3 databases
    """

    def __init__(self, db):
        """
        :param db: The filename of the database
        :type db: str
        """
        self.conn = sq.connect(db)
        self.cur = self.conn.cursor()

    def __del__(self):
        """
        Close the connection when the object is deleted
        :return:
        """
        self.conn.close()

    def commit(self):
        """
        Commit all changes
        """
        self.conn.commit()

    def close(self):
        """
        Close the connection
        """
        self.conn.close()

    def commitAndClose(self):
        """
        Commit all changes and close the connection
        """
        self.commit()
        self.close()

    def get_cursor(self, cursorClass=None):
        """
        Get a cursor to the database
        :return: A cursor to the database
        """
        if cursorClass:
            return self.conn.cursor(cursorClass)
        else:
            return self.conn.cursor()

    def execute(self, command, *args):
        """
        Execute a command
        :param command: The command to be executed
        :param args: The arguments to be passed to the command
        """
        self.cur.execute(command, *args)


class DbTable:

    """
    A class for handling sqlite3 tables which use an integer as the primary key
    """

    def __init__(self, db, tableName, autoCommit=False, **kwargs):
        """
        :param db: The name of the database
        :type db: DbConnection
        :param tableName: The name of the table
        :type tableName: str
        :param kwargs: Dictionary with keywords and values for the table
        """
        self.db = db
        self.cur = db.get_cursor()
        self.tableName = tableName
        self._autoCommit = autoCommit
        self.columnTypes = []
        command = 'CREATE TABLE IF NOT EXISTS ' + tableName + ' (id INTEGER PRIMARY KEY'
        for key in kwargs:
            command += ', ' + key + ' ' + kwargs[key]
            self.columnTypes.append(key)
        command += ')'
        self.db.execute(command)

    def insert(self, *args):
        """Insert a new line into the Database with arguments in the same order as they appear in the Database"""
        command = 'INSERT INTO ' + self.tableName + ' VALUES (NULL'
        for i in range(len(args)):
            command += ',?'
        command += ')'
        self.db.execute(command, args)
        self.autoCommit()

    def getAll(self):
        """Return all the data from the table"""
        self.cur.execute('SELECT * FROM ' + self.tableName)
        rows = self.cur.fetchall()
        return rows

    def search(self, **kwargs):
        """Return all rows where at least one of the column values matches the passed data"""
        command = 'SELECT * FROM ' + self.tableName + ' WHERE '
        pos = 0
        for arg in kwargs:
            if pos == 0:
                command = command + arg + '=? '
                pos += 1
            else:
                command = command + 'OR ' + arg + '=? '
        self.cur.execute(command, [kwargs[val] for val in kwargs])
        rows = self.cur.fetchall()
        return rows

    def delete(self, ident):
        """
        Delete the specified row from the table
        :param ident: the integer ID of the row to be deleted
        :param ident: int
        """
        self.db.execute('DELETE FROM ' + self.tableName + ' WHERE id=?', (ident,))
        self.autoCommit()

    def update(self, ident, **kwargs):
        """
        Update the specified row in the table with the passed values. Other columns untouched.
        :param ident: the integer ID of the row to be updated
        :param ident: int
        """
        command = 'UPDATE ' + self.tableName + ' SET '
        pos = 0
        for arg in kwargs:
            if pos == 0:
                command = command + arg + '=?'
                pos += 1
            else:
                command = command + ', ' + arg + '=?'
        command += 'WHERE id=?'
        args = [kwargs[val] for val in kwargs]
        args.append(ident)
        self.db.execute(command, args)
        self.autoCommit()

    def changeValue(self, newValue, key, oldValue):
        """Set all occurrences of oldValue in key column to newValue"""
        command = 'UPDATE ' + self.tableName + ' SET ' + key + '=? WHERE ' + key + '=?'
        self.db.execute(command, [newValue, oldValue])
        self.autoCommit()

    def replaceFrom(self, newValue, key, oldValue):
        """Replace oldValue within key column with newValue. Use changeValue to replace the value entirely"""
        command = 'UPDATE ' + self.tableName + ' SET ' + key + ' = REPLACE(' + key + ', ?, ?)'
        self.db.execute(command, [oldValue, newValue])
        self.autoCommit()

    def autoCommit(self):
        if self._autoCommit:
            self.db.commit()