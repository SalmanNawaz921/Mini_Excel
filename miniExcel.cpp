#include <iostream>
#include <fstream>
#include <sstream>
#include <vector>
#include <string>
#include <limits.h>
#include <conio.h>
#include <windows.h>
#include <iomanip>
#include <cstring>
#include <algorithm>

using namespace std;

enum Color
{
    Reset = 0,
    Green = 32,  // Green text
    Yellow = 33, // Yellow text
    Blue = 34,   // Blue text
    Red = 91     // Red text
};

template <typename T>
class MiniExcel
{
public:
    int actRow = 3;
    int actCol = 3;
    void gotoxy(int x, int y)
    {
        COORD coordinates;
        coordinates.X = x;
        coordinates.Y = y;
        SetConsoleCursorPosition(GetStdHandle(STD_OUTPUT_HANDLE), coordinates);
    }

    class Iterator
    {
        typename MiniExcel<T>::Cell *current;

    public:
        Iterator(typename MiniExcel<T>::Cell *current) : current(current)
        {
        }

        T &operator*()
        {
            if (current)
                return current->value;
            throw std::out_of_range("Iterator out of bounds");
        }

        Iterator &operator++()
        {
            if (current)
                current = current->down;
            return *this;
        }

        Iterator operator++(int)
        {
            if (current)
                current = current->right;
            return *this;
        }

        Iterator &operator--()
        {
            if (current)
                current = current->up;
            return *this;
        }
        Iterator &operator--(int)
        {
            if (current)
                current = current->left;
            return *this;
        }

        bool operator!=(const Iterator &other)
        {
            return (current != other.current);
        }
    };
    class Cell
    {

    public:
        T value;
        Cell *left;
        Cell *right;
        Cell *up;
        Cell *down;
        Color color;
        Cell(T value)
        {
            this->value = value;
            left = NULL;
            right = NULL;
            up = NULL;
            down = NULL;
        }
        void setColor(Color color)
        {
            this->color = color;
            cout << "\033[" << color << "m";
        }
        void resetColor()
        {
            cout << "\033[" << Reset << "m";
        }

    private:
        int x;
        int y;
    };

    MiniExcel(int rows = 5, int cols = 5)
    {
        start = nullptr;
        current = nullptr;
        this->rows = rows;
        this->cols = cols;
        initializeGrid(rows, cols);
        LoadFile("excelFile.csv");
    }

    void initializeGrid(int rows, int cols)
    {
        // Create the top-left cell.

        char a = 'A';
        Cell *current_cell = createCell(a + to_string(0));
        start = current_cell;
        current = current_cell;

        // Create the first row.

        for (int i = 0; i < cols - 1; i++)
        {
            Cell *newCell = createCell(char(++a) + to_string(0));
            current_cell->right = newCell;
            newCell->left = current_cell;
            current_cell = newCell;
        }

        // Create the remaining rows.

        Cell *firstCellInRow = current;
        for (int i = 1; i < rows; i++)
        {
            a = 'A';
            Cell *newRowCell = createCell(char(a++) + to_string(i));
            current = newRowCell;
            firstCellInRow->down = newRowCell;
            newRowCell->up = firstCellInRow;
            current_cell = newRowCell;

            // Create cells in the current row.

            for (int j = 1; j < cols; j++)
            {
                Cell *newCell = createCell(char(a++) + to_string(i));
                current_cell->right = newCell;
                newCell->left = current_cell;
                current_cell = newCell;
                firstCellInRow = firstCellInRow->right; // Move to the right in the first row.
                current_cell->up = firstCellInRow;      // Connect the cell to the cell above it.
                firstCellInRow->down = current_cell;    // Connect the cell above to the current cell.
            }
            firstCellInRow = newRowCell; // Move to the next row.
        }
    }

    void displayGrid(MiniExcel<T> &excel)
    {
        system("cls");
        Cell *currentCell = excel.getStart();
        int x = 3;
        int y = 3;
        this->printTheNumbers(getStart());
        while (currentCell)
        {
            Cell *row_Cell = currentCell;
            while (row_Cell)
            {
                printCell(row_Cell, x, y);
                row_Cell = row_Cell->right;
                x += 11;
            }
            x = 3;
            y += 4;

            cout
                << "\n\n\n";
            currentCell = currentCell->down;
        }
        getch();
    }
    void printTheNumbers(Cell *temp)
    {
        Cell *H = temp;
        int i = 1;
        int x = 8;
        int y = 5;
        char val = 'A';
        while (temp != nullptr || H != nullptr)
        {

            if (temp != nullptr)
            {
                gotoxy(x, 1);
                cout << "\033[31m" << i;
                i++;
                x += 11;
                temp = temp->right;
            }
            if (H != nullptr)
            {
                gotoxy(1, y);
                cout << "\033[33m" << val;
                val = val + 1;
                y += 4;
                H = H->down;
            }
        }
    }

    void printCell(Cell *temp, int x, int y)
    {
        gotoxy(x, y);
        temp->setColor(Blue);
        printCellBorder(x, y);
        gotoxy(x + 3, y + 2);
        cout << "\033[34m";
        printDataInCell(temp);
    }

    void printCellBorder(int &x, int &y)
    {
        cout
            << "----------";
        gotoxy(x, y + 1);
        cout << "|        |";
        gotoxy(x, y + 2);
        cout << "|        |";
        gotoxy(x, y + 3);
        cout << "|        |";
        gotoxy(x, y + 4);
        cout << "----------";
    }

    void printDataInCell(Cell *temp)
    {
        cout << temp->value;
    }

    void activeCell(Cell *temp, int x, int y)
    {
        gotoxy(x, y);
        temp->setColor(Yellow);
        cout << "##########";
        gotoxy(x, y + 1);
        cout << "#        #";
        gotoxy(x, y + 2);
        cout << "#        #";
        gotoxy(x, y + 3);
        cout << "#        #";
        gotoxy(x, y + 4);
        cout << "##########";
        gotoxy(x + 3, y + 2);
        cout << "\033[34m";
        printDataInCell(temp);
    }

    void MoveRight(Cell *temp)
    {
        if (temp->right != nullptr)
        {
            printCell(temp, actCol, actRow);
            temp = temp->right;
            setCurrent(temp);
            actCol += 11;
            activeCell(currentCell(), actCol, actRow);
        }
    }

    void MoveLeft(Cell *temp)
    {
        if (temp->left != nullptr)
        {
            printCell(temp, actCol, actRow);
            temp = temp->left;
            setCurrent(temp);
            actCol -= 11;
            activeCell(currentCell(), actCol, actRow);
        }
    }
    void MoveUp(Cell *temp)
    {
        if (temp->up != nullptr)
        {
            printCell(temp, actCol, actRow);
            temp = temp->up;
            setCurrent(temp);
            actRow -= 4;
            activeCell(currentCell(), actCol, actRow);
        }
    }
    void MoveDown(Cell *temp)
    {
        if (temp->down != nullptr)
        {
            printCell(temp, actCol, actRow);
            temp = temp->down;
            setCurrent(temp);
            actRow += 4;
            activeCell(currentCell(), actCol, actRow);
        }
    }

    void MoveCell(string direction)
    {
        Cell *temp = currentCell();
        transform(direction.begin(), direction.end(), direction.begin(), ::toupper);

        if (temp)
        {
            if (direction == "UP")
                MoveUp(temp);
            else if (direction == "DOWN")
                MoveDown(temp);
            else if (direction == "RIGHT")
                MoveRight(temp);
            else if (direction == "LEFT")
                MoveLeft(temp);
        }
    }

    Iterator begin()
    {
        return Iterator(start);
    }
    Iterator end()
    {
        return Iterator(nullptr);
    }

    Cell *currentCell()
    {
        return current;
    }

    // Function to insert a row below the current cell

    void InsertRowBelow(Cell *currentCell)
    {

        // Create a new row
        Cell *newRow = nullptr;
        Cell *currentRow = rowStartCell(currentCell);
        Cell *rowBelowCurrent = rowStartCell(currentCell)->down;
        // Initialize newRow with space character
        for (int i = 0; currentRow != nullptr; i++)
        {
            Cell *newRowCell = new Cell(" ");
            newRowCell->left = newRow;
            if (newRow != nullptr)
                newRow->right = newRowCell;

            newRowCell->up = currentRow;
            currentRow->down = newRowCell;
            newRow = newRowCell;
            currentRow = currentRow->right;
        }
        newRow = rowStartCell(newRow);
        for (int i = 0; rowBelowCurrent != nullptr; i++)
        {
            newRow->down = rowBelowCurrent;
            rowBelowCurrent->up = newRow;
            newRow = newRow->right;
            rowBelowCurrent = rowBelowCurrent->right;
        }
        rows++;
    }

    // Function to insert a row above the current cell
    void InsertRowAbove(Cell *currentCell)
    {
        // Create a new row
        Cell *newRow = nullptr;
        Cell *currentRow = rowStartCell(currentCell);
        Cell *rowUpCurrent = rowStartCell(currentCell)->up;
        // Initialize newRow with space character
        for (int i = 0; currentRow != nullptr; i++)
        {
            Cell *newRowCell = new Cell(" ");
            newRowCell->left = newRow;
            if (newRow != nullptr)
                newRow->right = newRowCell;

            newRowCell->down = currentRow;
            currentRow->up = newRowCell;
            newRow = newRowCell;
            currentRow = currentRow->right;
        }

        newRow = rowStartCell(newRow);
        if (!rowUpCurrent)
        {
            start = newRow;
        }

        for (int i = 0; rowUpCurrent != nullptr; i++)
        {
            newRow->up = rowUpCurrent;
            rowUpCurrent->down = newRow;
            rowUpCurrent = rowUpCurrent->right;
            newRow = newRow->right;
        }
        rows++;
    }

    // Function to insert a column to left of the current cell

    void InsertColumnToLeft(Cell *currentCell)
    {
        Cell *currentColumn = colStartCell(currentCell);
        Cell *columnToLeft = currentColumn->left;
        Cell *newCol = nullptr;
        for (int i = 0; currentColumn != nullptr; i++)
        {
            Cell *newCell = new Cell(" ");
            newCell->up = newCol;
            if (newCol != nullptr)
                newCol->down = newCell;
            newCell->right = currentColumn;
            currentColumn->left = newCell;
            newCol = newCell;
            currentColumn = currentColumn->down;
        }

        newCol = colStartCell(newCol);
        if (columnToLeft == nullptr)
            start = newCol;
        for (int i = 0; columnToLeft != nullptr; i++)
        {
            columnToLeft->right = newCol;
            newCol->left = columnToLeft;
            columnToLeft = columnToLeft->down;
            newCol = newCol->down;
        }
        cols++;
    }

    // Function to insert a column to right of the current cell

    void InsertColumnToRight(Cell *currentCell)
    {
        Cell *currentColumn = colStartCell(currentCell);
        Cell *columnToRight = currentColumn->right;
        Cell *newCol = nullptr;
        for (int i = 0; currentColumn != nullptr; i++)
        {
            Cell *newCell = new Cell(" ");
            newCell->up = newCol;
            if (newCol != nullptr)
                newCol->down = newCell;
            newCell->left = currentColumn;
            currentColumn->right = newCell;
            newCol = newCell;
            currentColumn = currentColumn->down;
        }

        newCol = colStartCell(newCol);

        for (int i = 0; columnToRight != nullptr; i++)
        {
            columnToRight->left = newCol;
            newCol->right = columnToRight;
            columnToRight = columnToRight->down;
            newCol = newCol->down;
        }
        cols++;
    }

    // Function to delete current row
    void DeleteRow()
    {
        Cell *currentRow = rowStartCell(current);
        Cell *rowAboveCurrent = currentRow->up;
        Cell *rowBelowCurrent = currentRow->down;
        if (!current->down)
            current = currentRow->up;
        if (!current->up)
        {
            current = currentRow->down;
            start = currentRow->down;
        }

        for (int i = 0; rowAboveCurrent != nullptr && rowBelowCurrent != nullptr; i++)
        {
            rowAboveCurrent->down = rowBelowCurrent;
            rowBelowCurrent->up = rowAboveCurrent;
            rowAboveCurrent = rowAboveCurrent->right;
            rowBelowCurrent = rowBelowCurrent->right;
        }

        if (!currentRow->up)
        {
            for (int i = 0; rowBelowCurrent != nullptr; i++)
            {
                rowBelowCurrent->up = nullptr;
                rowBelowCurrent = rowBelowCurrent->right;
            }
        }
        if (!currentRow->down)
        {
            for (int i = 0; rowAboveCurrent != nullptr; i++)
            {
                rowAboveCurrent->down = nullptr;
                rowAboveCurrent = rowAboveCurrent->right;
            }
        }
        rows--;
    }

    // Function to delete a current column

    void DeleteColumn()
    {
        Cell *currentCol = colStartCell(current);
        Cell *colToLeft = currentCol->left;
        Cell *colToRight = currentCol->right;
        if (!current->left)
            current = currentCol->right;
        if (!current->right)
            current = currentCol->left;

        for (int i = 0; colToLeft != nullptr && colToRight != nullptr; i++)
        {
            colToLeft->right = colToRight;
            colToRight->left = colToLeft;
            colToLeft = colToLeft->down;
            colToRight = colToRight->down;
        }

        if (!currentCol->left)
        {
            start = colToRight;
            while (colToRight != nullptr)
            {
                colToRight->left = nullptr;
                colToRight = colToRight->down;
            }
        }
        if (!currentCol->right)
        {
            while (colToLeft != nullptr)
            {
                colToLeft->right = nullptr;
                colToLeft = colToRight->down;
            }
        }
        cols--;
    }

    // Function to clear a column

    void ClearColumn()
    {
        Cell *currentCol = colStartCell(current);
        while (currentCol)
        {
            currentCol->value = " ";
            currentCol = currentCol->down;
        }
    }

    // Function to clear a row

    void ClearRow()
    {
        Cell *currentRow = rowStartCell(current);
        while (currentRow)
        {
            currentRow->value = " ";
            currentRow = currentRow->right;
        }
    }

    // Function to insert a new cell and shift the current cell to right

    void InsertCellByRightShift(Cell *current)
    {
        if (!current || !current->right)
        {
            return;
        }

        Cell *currentCell = current;
        Cell *tempCell = nullptr;
        string tempVal = " ";
        string nextVal = " ";
        while (currentCell)
        {
            nextVal = currentCell->value;
            currentCell->value = tempVal;
            tempVal = nextVal;
            tempCell = currentCell;
            currentCell = currentCell->right;
        }

        InsertColumnToRight(tempCell);
        tempCell->right->value = tempVal;
    }

    // Function to insert a new cell and shift the current cell to down
    void InsertCellByDownShift(Cell *current)
    {
        if (!current || !current->right)
        {
            return;
        }

        Cell *currentCell = current;
        Cell *tempCell = nullptr;
        string tempVal = " ";
        string nextVal = " ";
        while (currentCell)
        {
            nextVal = currentCell->value;
            currentCell->value = tempVal;
            tempVal = nextVal;
            tempCell = currentCell;
            currentCell = currentCell->down;
        }

        InsertRowBelow(tempCell);
        tempCell->down->value = tempVal;
    }
    // Function to delete the cell and shift the current cell to left

    void DeleteCellByLeftShift()
    {
        Cell *currentCell = current;
        while (currentCell->right)
        {
            Cell *next = currentCell->right;
            swap(currentCell->value, next->value);
            currentCell = next;
        }
        if (!currentCell->right && currentCell->left)
        {
            currentCell->left->right = nullptr;
        }

        if (currentCell->up)
        {
            currentCell->up->down = nullptr;
        }
        if (currentCell->down)
        {
            currentCell->down->up = nullptr;
        }
    }

    // Function to delete the cell and shift the current cell to up

    void DeleteCellByUpShift()
    {
        Cell *currentCell = current;
        while (currentCell->down)
        {
            Cell *next = currentCell->down;
            swap(currentCell->value, next->value);
            currentCell = next;
        }

        if (!currentCell->down && currentCell->up)
        {
            currentCell->up->down = nullptr;
            if (currentCell->right && !currentCell->left)
            {
                Cell *temp = currentCell->right;
                Cell *upCell = currentCell->up;
                while (temp && upCell)
                {
                    upCell->down = temp;
                    temp->up = upCell;
                    temp = temp->right;
                    upCell = upCell->right;
                }
            }
        }
        if (currentCell->left && currentCell->right)
        {
            currentCell->left->right = currentCell->right;
        }
        if (!current->left && !current->down && current->right)
        {
            current = current->right;
        }
    }

    // Operations Section

    // Function to get range sum of row/column
    double GetRangeSum(Cell *rangeStart, Cell *rangeEnd)
    {
        bool column = whetherRoworColumn(rangeStart, rangeEnd);
        double sum = 0;
        if (!column)
        {
            Cell *temp = rangeStart;
            while (temp != rangeEnd->right)
            {
                // As the value is mix of string and integer so (stoi) will help to distinguish between them
                try
                {
                    double cellValue = stod(temp->value);
                    sum += cellValue;
                }
                catch (const invalid_argument &e)
                {
                }
                temp = temp->right;
            }
        }
        else
        {
            Cell *temp = rangeStart;
            while (temp != rangeEnd->down)
            {
                try
                {
                    double cellValue = stod(temp->value);
                    sum += cellValue;
                }
                catch (const invalid_argument &e)
                {
                }
                temp = temp->down;
            }
        }
        return sum;
    }

    // Function to get average of row/column

    double GetRangeAverage(Cell *rangeStart, Cell *rangeEnd)
    {
        bool column = whetherRoworColumn(rangeStart, rangeEnd);
        double sum = 0;
        double nos = 0;
        if (!column)
        {
            Cell *temp = rangeStart;
            while (temp != rangeEnd->right)
            {
                try
                {
                    double cellValue = stod(temp->value);
                    sum += cellValue;
                    nos++;
                }
                catch (const invalid_argument &e)
                {
                }
                temp = temp->right;
            }
        }
        else
        {
            Cell *temp = rangeStart;
            while (temp != rangeEnd->down)
            {
                try
                {
                    double cellValue = stod(temp->value);
                    sum += cellValue;
                    nos++;
                }
                catch (const invalid_argument &e)
                {
                }
                temp = temp->down;
            }
        }
        return sum / nos;
    }

    // Function to find minimum number from a particular range
    double GetRangeMin(Cell *rangeStart, Cell *rangeEnd)
    {
        bool column = whetherRoworColumn(rangeStart, rangeEnd);
        double minNo = INT_MAX;
        if (!column)
        {
            Cell *temp = rangeStart;
            while (temp != rangeEnd->right)
            {
                if (temp->right && temp)
                {
                    try
                    {
                        double cellValue = stod(temp->value);
                        minNo = min(minNo, cellValue);
                    }
                    catch (const invalid_argument &e)
                    {
                        // Handle non-integer cell values if necessary
                    }
                }
                temp = temp->right;
            }
        }
        else
        {
            Cell *temp = rangeStart;
            while (temp != rangeEnd->down)
            {
                if (temp->down && temp)
                {
                    try
                    {
                        double cellValue = stod(temp->value);
                        minNo = min(minNo, cellValue);
                    }
                    catch (const invalid_argument &e)
                    {
                        // Handle non-integer cell values if necessary
                    }
                }
                temp = temp->down;
            }
        }
        return minNo;
    }
    // Function to find maximum number from a particular range
    double GetRangeMax(Cell *rangeStart, Cell *rangeEnd)
    {
        bool column = whetherRoworColumn(rangeStart, rangeEnd);
        double maxNo = INT_MIN;
        if (!column)
        {
            Cell *temp = rangeStart;
            while (temp != rangeEnd->right)
            {
                if (temp->right && temp)
                    try
                    {
                        double cellValue = stod(temp->value);
                        maxNo = max(maxNo, cellValue);
                    }
                    catch (const invalid_argument &e)
                    {
                        return 0.0;
                    }
                temp = temp->right;
            }
        }
        else
        {
            Cell *temp = rangeStart;
            while (temp != rangeEnd->down)
            {
                if (temp->down && temp)
                    try
                    {
                        double cellValue = stod(temp->value);
                        maxNo = max(maxNo, cellValue);
                    }
                    catch (const invalid_argument &e)
                    {
                        // Handle non-integer cell values if necessary
                    }
                temp = temp->down;
            }
        }
        return maxNo;
    }

    // Function to find counting number from a particular range
    int GetRangeCount(Cell *rangeStart, Cell *rangeEnd)
    {
        bool column = whetherRoworColumn(rangeStart, rangeEnd);
        int count = 0;
        if (!column)
        {
            Cell *temp = rangeStart;
            while (temp != rangeEnd->right)
            {
                count++;
                temp = temp->right;
            }
        }
        else
        {
            Cell *temp = rangeStart;
            while (temp != rangeEnd->down)
            {
                count++;
                temp = temp->down;
            }
        }
        return count;
    }

    // Function to copy the contents from particular range

    vector<T> Copy(Cell *startRange, Cell *endRange)
    {
        vector<T> new_range;
        bool column = whetherRoworColumn(startRange, endRange);
        if (!column)
        {
            Cell *temp = startRange;
            while (temp != endRange->right)
            {
                new_range.push_back(temp->value);
                temp = temp->right;
            }
        }
        else
        {
            Cell *temp = startRange;
            while (temp != endRange->down)
            {
                new_range.push_back(temp->value);
                temp = temp->down;
            }
        }
        return new_range;
    }
    // Function to cut the contents from particular range

    vector<T> Cut(Cell *startRange, Cell *endRange)
    {
        vector<T> new_range;
        bool column = whetherRoworColumn(startRange, endRange);
        if (!column)
        {
            Cell *temp = startRange;
            while (temp != endRange->right)
            {
                new_range.push_back(temp->value);
                temp->value = " ";
                temp = temp->right;
            }
        }
        else
        {
            Cell *temp = startRange;
            while (temp != endRange->down)
            {
                new_range.push_back(temp->value);
                temp->value = " ";
                temp = temp->down;
            }
        }
        return new_range;
    }

    // Function to paste the data of cut/copied cells

    void Paste(vector<T> &data)
    {
        int dataIndex = 0; // Index to access the data vector
        Cell *currentCell = current;

        if (data.size() <= rows)
        {
            currentCell = rowStartCell(current);
            while (dataIndex < data.size())
            {
                currentCell->value = data[dataIndex];
                dataIndex++;
                if (dataIndex < data.size())
                {
                    if (currentCell->right)
                    {
                        currentCell = currentCell->right;
                    }
                    else
                    {
                        // setCurrent(temp);
                        Cell *newCell = new Cell(" ");
                        currentCell->right = newCell;
                        newCell->left = currentCell;
                        currentCell = newCell;
                    }
                }
            }
        }
        else
        {
            currentCell = colStartCell(current);
            while (dataIndex < data.size())
            {
                if (dataIndex < data.size())

                {
                    currentCell->value = data[dataIndex];
                    dataIndex++;
                    if (currentCell->down)
                    {
                        currentCell = currentCell->down;
                    }
                    else
                    {
                        Cell *newCell = new Cell(" ");
                        currentCell->down = newCell;
                        newCell->up = currentCell;
                        currentCell = newCell;
                    }
                }
            }
        }
    }

    // Function to get middle cell position

    Cell *getMidCell()
    {
        int midRow = rows / 2;
        int midCol = cols / 2;
        Cell *midCell = start;
        for (int i = 0; i < midRow; i++)
        {
            midCell = midCell->down;
        }
        for (int i = 0; i < midCol; i++)
        {
            midCell = midCell->right;
        }
        return midCell;
    }

    // Function to get any cell position

    Cell *getCell(int rowIndex, int columnIndex)
    {
        Cell *startCell = start;

        for (int i = 0; i < rowIndex; i++)
        {
            startCell = startCell->down;
        }
        for (int j = 0; j < columnIndex; j++)
        {
            startCell = startCell->right;
        }

        return startCell;
    }

    // Function to set current cellc

    void setCurrent(Cell *current)
    {
        this->current = current;
    }

    // Function to get rows

    int getRows()
    {
        return rows;
    }

    // Function to get columns
    int getCols()
    {
        return cols;
    }

    // Function to set rows

    void setRows(int row)
    {
        row = rows;
    }

    // Function to set columns

    void setCols(int col)
    {
        col = cols;
    }

    // Function to  store the grid
    void storeGrid(string filename)
    {
        ofstream grid(filename);
        if (grid.is_open())
        {
            // File opened successfully, you can proceed with writing to it.
            Cell *rowCell = start;
            while (rowCell)
            {
                Cell *colCell = rowCell;
                while (colCell)
                {
                    grid << colCell->value << ",";
                    colCell = colCell->right;
                }
                grid << "\n";
                rowCell = rowCell->down;
            }
            grid.close(); // Close the file when done.
        }
        else
        {
            // Handle the case where the file couldn't be opened.
            cerr << "Failed to open the file for writing." << endl;
        }
    }

    // Function to load grid
    bool LoadFile(string filename)
    {
        ifstream file(filename);

        if (file.is_open())
        {
            string line;
            Cell *currentRow = start;
            int row = 0;

            while (getline(file, line))
            {
                istringstream iss(line);
                string value;

                int col = 0;
                while (getline(iss, value, ','))
                {

                    Cell *currentcell = getCell(row, col);
                    if (currentcell != nullptr)
                    {
                        currentcell->value = value;
                    }

                    col++;
                    if (col >= cols && currentcell)
                    {
                        Cell *colCell = colStartCell(currentcell);
                        InsertColumnToRight(colCell);
                    }
                }

                row++;
                if (row >= rows)
                {
                    if (currentRow)
                    {
                        Cell *rowCell = rowStartCell(currentRow);
                        InsertRowBelow(rowCell);
                    }
                }

                currentRow = currentRow->down;
            }

            file.close();
            setCurrent(getCell(0, 0));
            return true;
        }
        return false;
    }

    // Function to clear entire grid

    void clearGrid()
    {
        Cell *currentCell = start;
        while (currentCell)
        {
            Cell *rowCell = currentCell;
            while (rowCell)
            {
                // Assuming you have a function to set the cell value to a default or empty state
                rowCell->value = T(); // This assigns a default-constructed value for type T
                rowCell = rowCell->right;
            }
            currentCell = currentCell->down;
        }
    }

    // Function to get the start cell

    Cell *getStart()
    {
        return start;
    }

private:
    Cell *start;
    Cell *current;
    int rows;
    int cols;

    //                                      ************************** Private Functions To Increase Productivity **************************

    // Function to get row's start cell

    Cell *rowStartCell(Cell *currentRow)
    {
        Cell *currentCell = currentRow;
        while (currentCell->left)
        {
            currentCell = currentCell->left;
        }
        return currentCell;
    }

    // Function to get column's start cell

    Cell *colStartCell(Cell *currentCol)
    {
        Cell *currentCell = currentCol;
        while (currentCell->up)
        {
            currentCell = currentCell->up;
        }
        return currentCell;
    }

    // Function to check whether the given range is of column or row

    bool whetherRoworColumn(Cell *rangeStart, Cell *rangeEnd)
    {
        Cell *temp = rangeStart;
        while (temp != nullptr)
        {
            if (temp == rangeEnd)
                return true;
            temp = temp->down;
        }
        return false;
    }

    void setCellValue(int rowIndex, int colIndex, T value)
    {
        Cell *cell = getCell(rowIndex, colIndex);
        cell->value = value;
    }

    Cell *createCell(T value)
    {
        Cell *newCell = new Cell(value);
        newCell->setColor(Yellow);
        return newCell;
    }
};

void printColorText(const std::string &text, int colorCode)
{
    cout << "\033[38;5;" << colorCode << "m" << text << "\033[0m";
}

void printLogo()
{
    vector<string> logoText = {
        "                      ##                      ##     @@@@@@@@@@@    ##                     ##    @@@@@@@@@@@           @@           @@          @@@@@@@@     @@@@@@@@@@@     @@                             ",
        "                      @@  %%              %%  @@         ##         @@ %%                  @@         ##                %%        %%           ##            ##              ##                         ",
        "                      ##    %%           %%   ##         ##         ##   %%                ##         @@                 @@     @@            @@             @@              @@                        ",
        "                      @@      %%        %%    @@         ##         @@     %%              @@         ##                   %%  %%            ##              ##              ##                   ",
        "                      ##        %%     %%     ##         ##         ##       %%            ##         @@                     @@             @@               @@              @@                     ",
        "                      @@          %%  %%      @@         ##         @@         %%          @@         ##                     %% %%          ##               ##@@@@@@        ##                                    ",
        "                      ##            %%        ##         ##         ##           %%        ##         @@                    %%   %%         @@               @@              @@                     ",
        "                      @@                      @@         ##         @@             %%      @@         ##                   @@      @@        ##              ##              ##                  ",
        "                      ##                      ##         ##         ##               %%    ##         @@                 %%         %%        @@             @@              @@                  ",
        "                      @@                      @@         ##         @@                 %%  @@         ##                %%            %%       ##            ##              ##                  ",
        "                      ##                      ##      @@@@@@@@@@    ##                     ##     @@@@@@@@@@           @@               @@       @@@@@@@     @@@@@@@@@@@     @@@@@@@@@@@@@@                                  "};

    // Color codes for a rainbow effect
    int colorCodes[] = {9, 11, 13, 14, 12, 10};

    for (int i = 0; i < logoText.size(); ++i)
    {
        printColorText(logoText[i], colorCodes[i % (sizeof(colorCodes) / sizeof(colorCodes[0]))]);
        cout << endl;
    }
}

// Helper function to check if a key is down
bool IsKeyDown(int key)
{
    return GetAsyncKeyState(key) & 0x8000;
}

int printMenu()
{
    int opt;
    cout << "\n\t\t\t\t\t\t\t\t\t 1. Start" << endl;
    cout << "\t\t\t\t\t\t\t\t\t 2. Instructions" << endl;
    cout << "\t\t\t\t\t\t\t\t\t 3. Exit" << endl;
    cout << "\t\t\t\t\t\t\t\t\t Your Option.... ";
    cin >> opt;
    return opt;
}
void instructionsMenu()
{
    cout << "\n\t\t\t\t\t\t\t\t\tInstructions!!!\t\t\n"
         << endl;
    cout << " \t\t\t\t\t\tPress UP/DOWN Keys --> To Navigate The Sheet" << endl;
    cout << " \t\t\t\t\t\tPress I --> To Input The Value in a Cell" << endl;
    cout << " \t\t\t\t\t\tPress B --> To Select Starting Cell For Copy/Cut/Sum etc" << endl;
    cout << " \t\t\t\t\t\tPress C --> To Select Ending Cell & Call Copy()" << endl;
    cout << " \t\t\t\t\t\tPress X --> To Select Ending Cell & Call Cut()" << endl;
    cout << " \t\t\t\t\t\tPress V --> To Paste The Copy/Cut Data" << endl;
    cout << " \t\t\t\t\t\tPress Ctrl+A --> To Add Row Above Current Cell" << endl;
    cout << " \t\t\t\t\t\tPress Ctrl+B --> To Add Row Below Current Cell" << endl;
    cout << " \t\t\t\t\t\tPress Ctrl+R --> To Add Column Right of Current Cell" << endl;
    cout << " \t\t\t\t\t\tPress Ctrl+L --> To Add Column Left of Current Cell" << endl;
    cout << " \t\t\t\t\t\tPress Shift+R --> To Add a Cell & Shift Current Cell To The Right" << endl;
    cout << " \t\t\t\t\t\tPress Shift+D --> To Add a Cell & Shift Current Cell To The Down" << endl;
    cout << " \t\t\t\t\t\tPress Shift+L --> To Delete a Cell & Shift Current Cell To The Left" << endl;
    cout << " \t\t\t\t\t\tPress Shift+U --> To Delete a Cell & Shift Current Cell To The Up" << endl;
    cout << " \t\t\t\t\t\tPress Alt+R --> To Delete a Row" << endl;
    cout << " \t\t\t\t\t\tPress Alt+C --> To Delete a Column" << endl;
    cout << " \t\t\t\t\t\tPress Alt+R --> To Delete a Row" << endl;
    cout << " \t\t\t\t\t\tPress Alt+C --> To Delete a Column" << endl;
    cout << " \t\t\t\t\t\tPress Alt+S --> To Get The Sum of Your Desired Data" << endl;
    cout << " \t\t\t\t\t\tPress Alt+A --> To Get The Average of Your Desired Data" << endl;
    cout << " \t\t\t\t\t\tPress Alt+M --> To Get The Maximum of Your Desired Data" << endl;
    cout << " \t\t\t\t\t\tPress Alt+N --> To Get The Minimum of Your Desired Data" << endl;
    cout << " \t\t\t\t\t\tPress Alt+C --> To Get The Count of Your Desired Data" << endl;
    cout << " \t\t\t\t\t\tPress Ctrl+S --> To Save The File" << endl;
    cout << " \t\t\t\t\t\tPress Ctrl+W --> To Save The File & Exit" << endl;
    cout << " \t\t\t\t\t\tPress Escape --> To Exit" << endl;
}

void excelMenu()
{
    MiniExcel<string> excel;
    vector<string> copyOrCut;
    excel.setCurrent(excel.getCell(0, 0));
    MiniExcel<string>::Cell *rangeStart = nullptr;
    MiniExcel<string>::Cell *rangeEnd = nullptr;
    excel.displayGrid(excel);
    while (true)
    {
        // Up key for moving up

        if (GetAsyncKeyState(VK_UP))
        {
            excel.MoveCell("UP");
        }

        // Down key for moving down

        else if (GetAsyncKeyState(VK_DOWN))
        {
            excel.MoveCell("DOWN");
        }

        // Left key for moving left

        else if (IsKeyDown(VK_LEFT))
        {
            excel.MoveCell("LEFT");
        }

        // Right key for moving right

        else if (IsKeyDown(VK_RIGHT))
        {
            excel.MoveCell("RIGHT");
        }

        // Ctrl + A For Inserting Row Above Current Cell

        else if (IsKeyDown(VK_CONTROL) && IsKeyDown('A'))
        {
            excel.InsertRowAbove(excel.currentCell());
            excel.displayGrid(excel);
            excel.setCurrent(excel.getStart());
            excel.actCol = 3;
            excel.actRow = 3;
        }

        // Ctrl + B For Inserting Row Below Current Cell

        else if (IsKeyDown(VK_CONTROL) && IsKeyDown('B'))
        {
            excel.InsertRowBelow(excel.currentCell());
            excel.displayGrid(excel);
            excel.setCurrent(excel.getStart());
            excel.actCol = 3;
            excel.actRow = 3;
        }

        // Ctrl + R For Inserting Column To Right of Current Cell

        else if (IsKeyDown(VK_CONTROL) && IsKeyDown('R'))
        {
            excel.InsertColumnToRight(excel.currentCell());
            excel.displayGrid(excel);
            excel.setCurrent(excel.getStart());
            excel.actCol = 3;
            excel.actRow = 3;
        }

        // Ctrl + L For Inserting Column To Left Current Cell

        else if (IsKeyDown(VK_CONTROL) && IsKeyDown('L'))
        {
            excel.InsertColumnToLeft(excel.currentCell());
            excel.displayGrid(excel);
            excel.setCurrent(excel.getStart());
            excel.actCol = 3;
            excel.actRow = 3;
        }

        // Shift + R For Inserting a Cell & Shifting Current Cell To Right

        else if (IsKeyDown(VK_SHIFT) && IsKeyDown('R'))
        {
            excel.InsertCellByRightShift(excel.currentCell());
            excel.displayGrid(excel);
            excel.setCurrent(excel.getStart());
            excel.actCol = 3;
            excel.actRow = 3;
        }

        // Shift + D For Inserting a Cell & Shifting Current Cell To Down

        else if (IsKeyDown(VK_SHIFT) && IsKeyDown('D'))
        {
            excel.InsertCellByDownShift(excel.currentCell());
            excel.setCurrent(excel.getStart());
            excel.displayGrid(excel);
            excel.actCol = 3;
            excel.actRow = 3;
        }

        // Shift + L For Deleteing a Cell & Shifting Current Cell To Left

        else if (IsKeyDown(VK_SHIFT) && IsKeyDown('L'))
        {
            excel.DeleteCellByLeftShift();
            excel.setCurrent(excel.getStart());
            excel.displayGrid(excel);
            excel.actCol = 3;
            excel.actRow = 3;
        }

        // Shift + U For Deleteing a Cell & Shifting Current Cell To Up

        else if (IsKeyDown(VK_SHIFT) && IsKeyDown('U'))
        {
            excel.DeleteCellByUpShift();
            excel.setCurrent(excel.getStart());
            excel.displayGrid(excel);
            excel.actCol = 3;
            excel.actRow = 3;
        }

        // Alt + C For Deleting a Column

        else if (IsKeyDown(VK_MENU) && IsKeyDown('C'))
        {
            excel.DeleteColumn();
            excel.setCurrent(excel.getStart());
            excel.displayGrid(excel);
            excel.actCol = 3;
            excel.actRow = 3;
        }

        // Alt + R For Deleting a Row

        else if (IsKeyDown(VK_MENU) && IsKeyDown('R'))
        {
            excel.DeleteRow();
            excel.setCurrent(excel.getStart());
            excel.displayGrid(excel);
            excel.actCol = 3;
            excel.actRow = 3;
        }

        // C + R For Clearing a Row

        else if (IsKeyDown('C') && IsKeyDown('R'))
        {
            excel.ClearRow();
            excel.displayGrid(excel);
            excel.setCurrent(excel.getStart());
            excel.actCol = 3;
            excel.actRow = 3;
        }

        // C + L For Clearing a Column

        else if (IsKeyDown('C') && IsKeyDown('L'))
        {
            excel.ClearColumn();
            excel.displayGrid(excel);
            excel.setCurrent(excel.getStart());
            excel.actCol = 3;
            excel.actRow = 3;
        }

        // I For Inserting a value into a cell

        else if (IsKeyDown('I'))
        {
            string val;
            cin >> val;
            if (val.length() < 4 && val.length() > 0)
                excel.currentCell()->value = val;
            excel.displayGrid(excel);
            cin.clear();
        }

        // B For Selecting RangeStart

        else if (IsKeyDown('B'))
        {
            rangeStart = excel.currentCell();
        }

        // C For Selecting RangeEnd & Copying From RangeStart To RangeEnd

        else if (IsKeyDown('C'))
        {
            rangeEnd = excel.currentCell();
            copyOrCut = excel.Copy(rangeStart, rangeEnd);
        }

        // X For Selecting RangeEnd & Cut From RangeStart To RangeEnd

        else if (IsKeyDown('X'))
        {
            rangeEnd = excel.currentCell();
            copyOrCut = excel.Cut(rangeStart, rangeEnd);
            excel.displayGrid(excel);
        }

        // V For Pasting From RangeStart To RangeEnd

        else if (IsKeyDown('V'))
        {
            excel.Paste(copyOrCut);
            excel.displayGrid(excel);
        }

        // Alt + S for sum of the desired data

        else if (IsKeyDown(VK_MENU) && IsKeyDown('S'))
        {
            double sum = excel.GetRangeSum(rangeStart, rangeEnd);
            excel.currentCell()->value = to_string(sum);
            excel.displayGrid(excel);
        }

        // Alt + A for average of the desired data

        else if (IsKeyDown(VK_MENU) && IsKeyDown('A'))
        {
            double average = excel.GetRangeAverage(rangeStart, rangeEnd);
            excel.currentCell()->value = to_string(average);
            excel.displayGrid(excel);
        }

        // Alt + M for maximum of the desired data

        else if (IsKeyDown(VK_MENU) && IsKeyDown('M'))
        {
            double maxNo = excel.GetRangeMax(rangeStart, rangeEnd);
            excel.currentCell()->value = to_string(maxNo);
            excel.displayGrid(excel);
        }

        // Alt + M for minimum of the desired data

        else if (IsKeyDown(VK_MENU) && IsKeyDown('N'))
        {
            double minNo = excel.GetRangeMin(rangeStart, rangeEnd);
            excel.currentCell()->value = to_string(minNo);
            excel.displayGrid(excel);
        }

        // Ctrl + C for counting of the desired data

        else if (IsKeyDown(VK_MENU) && IsKeyDown('C'))
        {
            int count = excel.GetRangeCount(rangeStart, rangeEnd);
            excel.currentCell()->value = to_string(count);
            excel.displayGrid(excel);
        }

        // Ctrl + S for saving/storing the sheet as .csv file

        else if (IsKeyDown(VK_CONTROL) && IsKeyDown('S'))
        {
            excel.storeGrid("excelFile.csv");
        }
        // Ctrl + S for saving/storing the sheet as .csv file & exit

        else if (IsKeyDown(VK_CONTROL) && IsKeyDown('W'))
        {
            excel.storeGrid("excelFile.csv");
            break;
        }

        // Escape for exiting the loop

        else if (IsKeyDown(VK_ESCAPE))
        {
            break;
        }

        // Adjust the sleep time to control the frequency of key checks and updates.
        Sleep(100);
    }
}

int main()
{
    int opt = 0;
    system("cls");
    while (opt != 3)
    {
        printLogo();
        opt = printMenu();
        if (opt == 1)
        {

            system("cls");
            // printLogo();
            excelMenu();
        }
        else if (opt == 2)
        {
            system("cls");
            printLogo();
            instructionsMenu();
        }
        cout << "Press any key to continue...";
        getch();
        system("cls");
    }
}