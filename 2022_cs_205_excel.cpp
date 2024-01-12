#include <iostream>
#include <windows.h>
#include <conio.h>
#include <iomanip>
#include <fstream>
#include <string>
#include <vector>
#include <limits.h>
#include <cmath>
#include <cctype>
#include <sstream>
using namespace std;
class Excel
{
	class Cell
	{
	public:
		friend class Excel;
		string data;
		Cell *up;
		Cell *down;
		Cell *left;
		Cell *right;

		Cell(string input = " ", Cell *Left = nullptr, Cell *Right = nullptr, Cell *Up = nullptr, Cell *Down = nullptr)
		{
			data = input;
			left = Left;
			right = Right;
			up = Up;
			down = Down;
		}
	};
	Cell *head;
	Cell *current;
	Cell *RangeStart;
	vector<vector<string>> Clipboard;
	int row_size, col_size;
	int c_row, c_col;
	int Cell_width = 8;
	int Cell_height = 4;
	void color(int c)
	{
		HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
		SetConsoleTextAttribute(hConsole, c);
		if (c > 255)
		{
			c = 1;
		}
	}
	void gotoRowCol(int cinput, int rinput)
	{
		COORD scrn;
		HANDLE hOutput = GetStdHandle(STD_OUTPUT_HANDLE);
		scrn.X = rinput;
		scrn.Y = cinput;
		SetConsoleCursorPosition(hOutput, scrn);
	}

	void Print_cell(int col, int row, int colour)
	{
		color(colour);
		char c = '#';
		int startCol = col * Cell_height;
		int startRow = row * Cell_width;

		for (int i = 0; i <= Cell_height; i++)
		{
			gotoRowCol(startCol + i, startRow);
			cout << c;
			gotoRowCol(startCol + i, startRow + Cell_width);
			cout << c;
		}

		for (int i = 0; i <= Cell_width; i++)
		{
			gotoRowCol(startCol, startRow + i);
			cout << c;
			gotoRowCol(startCol + Cell_height, startRow + i);
			cout << c;
		}
	}
	void Print_cell_data(int row, int col, Cell *cell, int colour)
	{
		color(colour);
		gotoRowCol((Cell_height * row) + Cell_height / 2, (col * Cell_width) + Cell_width / 2);
		cout << "    ";
		gotoRowCol((Cell_height * row) + Cell_height / 2, (col * Cell_width) + Cell_width / 2);
		cout << cell->data;
	}
	public:
	Excel()
	{
		head = nullptr;
		current = nullptr;
		row_size = col_size = 5; 
		c_row = c_col = 0;
		head = Form_row();
		current = head;
		for (int i = 0; i < row_size - 1; i++)
		{
			InsertRowBelow();
			row_size--;
			current = current->down;
		}
		current = head;
		Print_Grid();
		Print_Data();
	}
	Cell *Form_row()
	{
		Cell *temp = new Cell();
		Cell *curr = temp;
		for (int i = 0; i < col_size - 1; i++)
		{
			Cell *temp2 = new Cell();
			temp->right = temp2;
			temp2->left = temp;
			temp = temp2;
		}
		return curr;
	}
	Cell *Form_col()
	{
		Cell *temp = new Cell();
		Cell *curr = temp;
		for (int i = 0; i < row_size - 1; i++)
		{
			Cell *temp2 = new Cell();
			temp->down = temp2;
			temp2->up = temp;
			temp = temp2;
		}
		return curr;
	}
	void InsertColRight()
	{
		Cell *newColumn = Form_col();
		Cell *topmost = current;
		while (topmost->up != nullptr)
		{
			topmost = topmost->up;
		}

		while (topmost != nullptr)
		{
			if (topmost->right != nullptr)
			{
				topmost->right->left = newColumn;
				newColumn->right = topmost->right;
			}
			topmost->right = newColumn;
			newColumn->left = topmost;
			topmost = topmost->down;
			newColumn = newColumn->down;
		}
		col_size++;
	}

	void InsertColLeft()
	{
		Cell *newColumn = Form_col(); 
		Cell *topmost = current;
		while (topmost->up != nullptr)
		{
			topmost = topmost->up;
		}
		if (topmost == head)
		{
			head = newColumn; 
		}
		while (topmost != nullptr)
		{
			if (topmost->left != nullptr)
			{
				topmost->left->right = newColumn;
				newColumn->left = topmost->left;
			}
			newColumn->right = topmost;
			topmost->left = newColumn;
			topmost = topmost->down;
			newColumn = newColumn->down;
		}
		col_size++; 
		Print_Grid(); 
		Print_Data(); 
	}

	void InsertRowBelow()
	{
		Cell *temp = Form_row();
		Cell *temp2 = current;
		while (temp2->left != nullptr)
		{
			temp2 = temp2->left;
		}
		if (temp2->down == nullptr)
		{
			while (temp2 != nullptr)
			{
				temp2->down = temp;
				temp->up = temp2;
				temp = temp->right;
				temp2 = temp2->right;
			}
		}
		else
		{
			while (temp2 != nullptr)
			{
				temp->down = temp2->down;
				temp2->down = temp;
				temp->up = temp2;
				temp->down->up = temp;
				temp = temp->right;
				temp2 = temp2->right;
			}
		}
		row_size++;
	}
	void InsertRowAbove()
	{
		Cell *temp =Form_row();
		Cell *temp2 = current;
		while (temp2->left != nullptr)
		{
			temp2 = temp2->left;
		}
		if (temp2 == head)
		{
			head = temp;
		}
		if (temp2->up == nullptr)
		{
			while (temp2 != nullptr)
			{
				temp2->up = temp;
				temp->down = temp2;
				temp = temp->right;
				temp2 = temp2->right;
			}
		}
		else
		{
			while (temp2 != nullptr)
			{
				temp->up = temp2->up;
				temp2->up = temp;
				temp->down = temp2;
				temp->up->down = temp;
				temp = temp->right;
				temp2 = temp2->right;
			}
		}
		row_size++;
		Print_Grid();
		Print_Data();
	}

	
		void InsertCellByRightShift()
		{
			Cell *temp = current;
			while (current->right != nullptr)
			{
				current = current->right;
			}
			InsertColRight();
			current = current->right;
			while (current != temp)
			{
				current->data = current->left->data;
				current = current->left;
			}
			current->data = "   ";
		}
	void InsertCellByDownShift()
	{
		Cell *temp = current;
		while (current->down != nullptr)
		{
			current = current->down;
		}
		InsertRowBelow();
		current = current->down;
		while (current != temp)
		{
			current->data = current->up->data;
			current = current->up;
		}
		current->data = " ";
	}
void DeleteCellByLeftShift()
{
    Cell *temp = current;
    Cell *prev = nullptr;

    while (temp != nullptr)
    {
        if (temp->right != nullptr)
        {
            temp->data = temp->right->data;
        }
        else
        {
            temp->data = "    ";
        }
        
        prev = temp;
        temp = temp->right;
    }
}

	void DeleteCellByUpShift()
{
    Cell *temp = current;
    Cell *prev = nullptr;

    while (temp != nullptr)
    {
        if (temp->down != nullptr)
        {
            temp->data = temp->down->data;
        }
        else
        {
            temp->data = " ";
        }
        
        prev = temp;
        temp = temp->down;
    }
}
	void Delete_Col()
	{
		Cell *temp = current;
		Cell *deleting_cell;
		if (col_size <= 1)
			return;
		while (temp->up != nullptr)
		{
			temp = temp->up;
		}
		if (temp == head)
		{
			head = temp->right;
		}
		if (temp->right == nullptr)
		{
			c_col--;
			current = current->left;
			while (temp != nullptr)
			{
				deleting_cell = temp;
				temp->left->right = nullptr;
				temp = temp->down;
				delete deleting_cell;
			}
		}
		else if (temp->left == nullptr)
		{
			current = current->right;
			while (temp != nullptr)
			{
				deleting_cell = temp;
				temp->right->left = nullptr;

				temp = temp->down;
				delete deleting_cell;
			}
		}
		else
		{
			c_col--;
			current = current->left;
			while (temp != nullptr)
			{
				deleting_cell = temp;
				temp->left->right = temp->right;
				temp->right->left = temp->left;

				temp = temp->down;
				delete deleting_cell;
			}
		}
		col_size--;
	}
	void Delete_Row()
	{
		Cell *temp = current;
		Cell *deleting_cell;
		if (row_size <= 1)
			return;
		while (temp->left != nullptr)
		{
			temp = temp->left;
		}
		if (temp == head)
		{
			head = temp->down;
		}
		if (temp->down == nullptr)
		{
			c_row--;
			current = current->up;
			while (temp != nullptr)
			{
				deleting_cell = temp;
				temp->up->down = nullptr;
				temp = temp->right;
				delete deleting_cell;
			}
		}
		else if (temp->up == nullptr)
		{
			current = current->down;
			while (temp != nullptr)
			{
				deleting_cell = temp;
				temp->down->up = nullptr;

				temp = temp->right;
				delete deleting_cell;
			}
		}
		else
		{
			c_row--;
			current = current->up;
			while (temp != nullptr)
			{
				deleting_cell= temp;
				temp->down->up = temp->up;
				temp->up->down = temp->down;

				temp = temp->right;
				delete deleting_cell;
			}
		}
		row_size--;
	}
	void Clear_Col()
	{
		Cell *temp = current;
		while (temp->up != nullptr)
		{
			temp = temp->up;
		}
		while (temp != nullptr)
		{
			temp->data = "    ";
			temp = temp->down;
		}
	}
	void Clear_Row()
	{
		Cell *temp = current;
		while (temp->left != nullptr)
		{
			temp = temp->left;
		}
		while (temp != nullptr)
		{
			temp->data = "    ";
			temp = temp->right;
		}
	}
	void Print_Grid()
	{
		for (int ri = 0; ri < row_size; ri++)
		{
			for (int ci = 0; ci < col_size; ci++)
			{
				Print_cell(ri, ci, 9);
			}
		}
	}
	void ClearRow_Col()
	{
		char c = _getch();
		if (c == 99)
	    {
			Clear_Col();
			Print_Data(); // Clear column ke lia(c)
		}
		else if (c == 114)
		{
			Clear_Row(); // Clear row ke lia(r)
			Print_Data();
		}
	}
	void Print_Data()
	{
		Cell *temp = head;
		for (int ri = 0; ri < row_size; ri++)
		{
			Cell *temp2 = temp;
			for (int ci = 0; ci < col_size; ci++)
			{
				Print_cell_data(ri, ci, temp, 9);
				temp = temp->right;
			}

			temp = temp2->down;
		}
	}

	bool check_string_digit(Cell *temp)
{
    if (temp->data.empty())
        return false;

    for (int i = 0; i < temp->data.length(); ++i)
    {
        if (i == 0 && (temp->data[i] == '-' || temp->data[i] == '+'))
        {
            continue;
        }

        if (!isdigit(temp->data[i]))
        {
            return false;
        }
    }
    return true;
}

	int Calculate_sum(string &startRow, string &startCol, string &endRow, string &endCol)
{
	if(!is_integer(startRow) || !is_integer(startCol ) || !is_integer(endRow) || !is_integer(endCol))
	{
		return INT_MIN;
	}
	if (stoi(startRow) < 0 || stoi(startCol) < 0 || stoi(endRow) >= row_size || stoi(endCol) >= col_size ||
    stoi(startRow) > stoi(endRow) || stoi(startCol) > stoi(endCol))
	{
		return INT_MIN;
	}
    Cell *temp = head;
    int sum = 0;
    for (int i = 0; i < stoi(startRow); ++i)
    {
        temp = temp->down;
    }
    for (int i = 0; i < stoi(startCol); ++i)
    {
        temp = temp->right;
    }
    for (int row = stoi(startRow); row <= stoi(endRow); ++row)
    {
        Cell *temp2 = temp;
        for (int col = stoi(startCol); col <= stoi(endCol); ++col)
        {
            if (check_string_digit(temp))
            {
                sum += stoi(temp->data);
            }
            temp = temp->right;
        }

        temp = temp2->down;
    }
    return sum;
}

	int Calculate_average(string &startRow, string &startCol, string &endRow, string &endCol)
    {
    if(!is_integer(startRow) || !is_integer(startCol ) || !is_integer(endRow) || !is_integer(endCol))
	{
		return INT_MIN;
	}
	if (stoi(startRow) < 0 || stoi(startCol) < 0 || stoi(endRow) >= row_size || stoi(endCol) >= col_size ||
    stoi(startRow) > stoi(endRow) || stoi(startCol) > stoi(endCol))
	{
		return INT_MIN;
	}
    Cell *temp = head; 
    int sum = 0;
    int count = 0;
    for (int i = 0; i < stoi(startRow); ++i)
    {
        temp = temp->down;
    }
    for (int i = 0; i < stoi(startCol); ++i)
    {
        temp = temp->right;
    }
    for (int row = stoi(startRow); row <= stoi(endRow); ++row)
    {
        Cell *temp2 = temp;
        for (int col = stoi(startCol); col <= stoi(endCol); ++col)
  {
            if (check_string_digit(temp))
            {
                sum += std::stoi(temp->data);
                count++;
            }
            temp = temp->right;
        }

        temp = temp2->down;
    }

    if (count > 0)
    {
        return (sum / count);
    }
    else
    {
        return 0;
    }
}

	int Calculate_Count(string &startRow, string &startCol, string &endRow, string &endCol)
{
	if(!is_integer(startRow) || !is_integer(startCol ) || !is_integer(endRow) || !is_integer(endCol))
	{
		return INT_MIN;
	}
	if (stoi(startRow) < 0 || stoi(startCol) < 0 || stoi(endRow) >= row_size || stoi(endCol) >= col_size ||
    stoi(startRow) > stoi(endRow) || stoi(startCol) > stoi(endCol))
	{
		return INT_MIN;
	}
    Cell *temp = head; 
    int count = 0;
    for (int i = 0; i < stoi(startRow); ++i)
    {
        temp = temp->down;
    }
    for (int i = 0; i < stoi(startCol); ++i)
    {
        temp = temp->right;
    }
    for (int row = stoi(startRow); row <= stoi(endRow); ++row)
    {
        Cell *temp2 = temp;
        for (int col = stoi(startCol); col <= stoi(endCol); ++col)
        {
            if (temp->data != " ")
                count++;
            temp = temp->right;
        }

        temp = temp2->down;
    }

    return count;
}
	int Calculate_Max(string &startRow, string &startCol, string &endRow, string &endCol)
{
	if(!is_integer(startRow) || !is_integer(startCol ) || !is_integer(endRow) || !is_integer(endCol))
	{
		return INT_MIN;
	}
	if (stoi(startRow) < 0 || stoi(startCol) < 0 || stoi(endRow) >= row_size || stoi(endCol) >= col_size ||
    stoi(startRow) > stoi(endRow) || stoi(startCol) > stoi(endCol))
	{
		return INT_MIN;
	}
    Cell *temp = head; 
    int Max = INT_MIN;
    for (int i = 0; i < stoi(startRow); ++i)
    {
        temp = temp->down;
    }
    for (int i = 0; i < stoi(startCol); ++i)
    {
        temp = temp->right;
    }
    for (int row = stoi(startRow); row <= stoi(endRow); ++row)
    {
        Cell *temp2 = temp;
        for (int col = stoi(startCol); col <= stoi(endCol); ++col)
        {
            if (check_string_digit(temp))
            {
                if (Max < stoi(temp->data))
                {
                    Max = stoi(temp->data);
                }
            }
            temp = temp->right;
        }

        temp = temp2->down;
    }

    return Max;
}

	int Calculate_Min(string &startRow, string &startCol, string &endRow, string &endCol)
{
	if(!is_integer(startRow) || !is_integer(startCol ) || !is_integer(endRow) || !is_integer(endCol))
	{
		return INT_MIN;
	}
	if (stoi(startRow) < 0 || stoi(startCol) < 0 || stoi(endRow) >= row_size || stoi(endCol) >= col_size ||
    stoi(startRow) > stoi(endRow) || stoi(startCol) > stoi(endCol))
	{
		return INT_MIN;
	}
    Cell *temp = head; 
    int Min = INT_MAX;
    for (int i = 0; i < stoi(startRow); ++i)
    {
        temp = temp->down;
    }
    for (int i = 0; i < stoi(startCol); ++i)
    {
        temp = temp->right;
    }
    for (int row = stoi(startRow); row <= stoi(endRow); ++row)
    {
        Cell *temp2 = temp;
    for (int col = stoi(startCol); col <= stoi(endCol); ++col)
        {
            if (check_string_digit(temp))
            {
                if (Min > std::stoi(temp->data))
                {
                    Min = std::stoi(temp->data);
                }
            }
            temp = temp->right;
        }

        temp = temp2->down;
    }

    return Min;
}
string Copy(string &startRow, string &startCol, string &endRow, string &endCol)
{
    if(!is_integer(startRow) || !is_integer(startCol ) || !is_integer(endRow) || !is_integer(endCol))
	{
		return "invalid";
	}
	if (stoi(startRow) < 0 || stoi(startCol) < 0 || stoi(endRow) >= row_size || stoi(endCol) >= col_size ||
    stoi(startRow) > stoi(endRow) || stoi(startCol) > stoi(endCol))
	{
		return "invalid";
	}
    Clipboard.clear();
    Cell *temp = head;
    for (int i = 0; i < stoi(startRow); ++i)
    {
        temp = temp->down;
    }
    for (int i = 0; i < stoi(startCol); ++i)
    {
        temp = temp->right;
    }
    for (int row = stoi(startRow); row <= stoi(endRow); ++row)
    {
        vector<string> clipRow;
        Cell *temp2 = temp;

        for (int col = stoi(startCol); col <= stoi(endCol); ++col)
        {
            clipRow.push_back(temp2->data);
            temp2 = temp2->right;
        }
        Clipboard.push_back(clipRow);
        temp = temp->down;
    }
	return "copied successfully";
}

	string Cut(string &startRow, string &startCol, string &endRow, string &endCol)
{
    if(!is_integer(startRow) || !is_integer(startCol ) || !is_integer(endRow) || !is_integer(endCol))
	{
		return "invalid" ;
	}
	if (stoi(startRow) < 0 || stoi(startCol) < 0 || stoi(endRow) >= row_size || stoi(endCol) >= col_size ||
    stoi(startRow) > stoi(endRow) || stoi(startCol) > stoi(endCol))
	{
		return "invalid";
	}

    Clipboard.clear();
	this->Copy(startRow, startCol, endRow, endCol);
	Cell* temp = head;
	for (int i = 0; i < stoi(startRow); ++i)
	{
		temp = temp->down;
	}
	for (int i = 0; i < stoi(startCol); ++i)
	{
		temp = temp->right;
	}
	for (int row = stoi(startRow); row <= stoi(endRow); ++row)
	{
		
		Cell* temp2 = temp;

		for (int col = stoi(startCol); col <= stoi(endCol); ++col)
		{
			temp2->data = " ";
			temp2 = temp2->right;
		}
		temp = temp->down;
	}
	    return "Cut succesfully";
}

void Paste()
{
    Cell *temp = current;
    for (int ri = 0; ri < Clipboard.size(); ri++)
    {
        Cell *temp2 = current;
        for (int ci = 0; ci < Clipboard[0].size(); ci++)
        {
            current->data = Clipboard[ri][ci];
            
            if (ci < Clipboard[0].size() - 1)
            {
                if (current->right == nullptr)
				{
                    InsertColRight();
				}
                current = current->right;
            }
        }
        if (ri < Clipboard.size() - 1)
        {
            if (temp2->down == nullptr)
                InsertRowBelow();
            else
                current = temp2->down;
        }
    }

    current = temp;
}
	void Mathematic_operations(string &n1,string &n2,string &n3,string &n4)
	{
		string s;
		char c = getch();
		RangeStart = current;
		if (c == 115) // Sum ke lia s
		{
			input(n1,n2,n3,n4);
			s = to_string(Calculate_sum(n1,n2,n3,n4));
			if(stoi(s)!=INT_MIN)
			cout << "Sum = " << s << " ";
			else
			cout <<"invalid  ";
		}
		else if (c == 97) // Average ke lia a
		{
			input(n1,n2,n3,n4);
			s = to_string(Calculate_average(n1,n2,n3,n4));
			if(stoi(s)!=INT_MIN)
			cout << "Average = " << s << " ";
			else
			cout <<"invalid  ";
		}
		else if (c == 99) // Count ke lia c
		{
			input(n1,n2,n3,n4);
			s = to_string(Calculate_Count(n1,n2,n3,n4));
			if(stoi(s)!=INT_MIN)
			cout << "Count = " << s << " ";
			else
			cout <<"invalid  ";
		}
		else if (c == 109) // Max ke lia m
		{
			input(n1,n2,n3,n4);
			s = to_string(Calculate_Max(n1,n2,n3,n4));
			if(stoi(s)!=INT_MIN)
			cout << "Max = " << s << " ";
			else
			cout <<"invalid  ";
		}
		else if (c == 110) // Min ke lia n
		{
			input(n1,n2,n3,n4);
			s = to_string(Calculate_Min(n1,n2,n3,n4));
			if(stoi(s)!=INT_MIN)
			cout << "Min = " << s << " ";
			else
			cout <<"invalid  ";
		}
		system("pause");
		gotoRowCol(Cell_height * row_size + 10+8, 0);
		cout << "                                                          ";
		current = RangeStart;
		if(stoi(s)!=INT_MIN)
		current->data = s;
		Print_Grid();
		Print_Data();
	}

	void Shift_and_delete()
	{
		char c = _getch();
		if (c == 100) // right cell shift ke lia d
		{

			InsertCellByRightShift();
			Print_Grid();
			Print_Data();
		}
		else if (c == 115) // down cell shift ke lia s
		{
			InsertCellByDownShift();
			Print_Grid();
			Print_Data();
		}
		else if (c == 97) // Delete shift left ke lia a
		{
			DeleteCellByLeftShift();
			Print_Grid();
			Print_Data();
		}
		else if (c == 119) // Delete shift up ke lia w
		{
			DeleteCellByUpShift();
			Print_Grid();
			Print_Data();
		}
		else if (c == 114) // delete row ke lia r
		{
			Delete_Row();
			system("cls");
			Print_Grid();
			Print_Data();
		}
		else if (c == 99) // delete col ke lia c
		{
			Delete_Col();
			system("cls");
			Print_Grid();
			Print_Data();
		}
	}

	void Cut_Copy_Paste(string &n1,string &n2,string &n3,string &n4)
	{
		string s;
		char c = _getch();
		RangeStart = current;
		if (c == 99) // Copy ke lia c
		{
			input(n1,n2,n3,n4);
			s=Copy(n1,n2,n3,n4);
			cout << s << endl;
		}
		else if (c == 112) // Paste ke lia p
		{
			Paste();
			gotoRowCol(Cell_height * row_size + 10+8, 0);
			cout << "pasted succesfully" << endl;
		}
		else if (c == 120) // Cut ke lia x
		{
			input(n1,n2,n3,n4);
			s=Cut(n1,n2,n3,n4);
			cout << s << endl;
		}
		current = RangeStart;
		Print_Grid();
		Print_Data();
		gotoRowCol(Cell_height * row_size + 10+9, 0);
		system("pause");
		gotoRowCol(Cell_height * row_size + 10+8, 0);
		cout << "                                                          ";
		gotoRowCol(Cell_height * row_size + 10+9, 0);
		cout << "                                                          ";
	}
	bool is_integer(string& s)
	{
		stringstream ss(s);
		int x;
		return (ss >> x) && (ss.eof());
	}
    void input(string &n1,string &n2,string &n3,string &n4)
   {
	gotoRowCol(Cell_height * row_size + 10, 0);
	cout << "enter your starting row:";
	cin >> n1;
	gotoRowCol(Cell_height * row_size + 10+2, 0);
	cout << "enter your starting col:";
	cin >> n2;
	gotoRowCol(Cell_height * row_size + 10+4, 0);
	cout << "enter your ending row:";
	cin >> n3;
	gotoRowCol(Cell_height * row_size + 10+6, 0);
	cout << "enter your ending col:";
	cin >> n4;
	gotoRowCol(Cell_height * row_size + 10, 0);
	cout <<"                             ";
	gotoRowCol(Cell_height * row_size + 10+2, 0);
	cout <<"                             ";
	gotoRowCol(Cell_height * row_size + 10+4, 0);
	cout <<"                             ";
	gotoRowCol(Cell_height * row_size + 10+6, 0);
	cout <<"                             ";
	gotoRowCol(Cell_height * row_size + 10+8, 0);
   }
	void Keyboard()
	{
	    string n1,n2,n3,n4;
		Print_cell(c_row, c_col, 5);
		Cell *temp = current;
		while (true)
		{
			char c = _getch();
			if (c == 100) // right ke lia d
			{
				if (current->right == nullptr)
				{
					InsertColRight();
					Print_Grid();
					Print_Data();
				}
				Print_cell(c_row, c_col, 9);
				Print_cell_data(c_row, c_col, current, 9);
				current = current->right;
				c_col++;
			}
			else if (c == 101) //e current ke right pr dalne ke lia
			{
				if (current->right != nullptr)
				{
					InsertColRight();
					Print_Grid();
					Print_Data();
				}
				Print_cell(c_row, c_col, 9); 
				Print_cell_data(c_row, c_col, current, 9);
				current = current->right;
				c_col++;
			}
			else if (c == 97) // left ke lia a
			{
				if (current->left == nullptr)
				{
					InsertColLeft();
					Print_Grid();
					Print_Data();
					c_col++;
				}
				Print_cell(c_row, c_col, 9); 
				Print_cell_data(c_row, c_col, current, 9);
				current = current->left;
				c_col--;
			}
			else if (c == 114) // r current ke left pr dalne ke lia
			{
				if (current->left != nullptr)
				{
					InsertColLeft();
					Print_Grid();
					Print_Data();
					c_col++;
				Print_cell(c_row, c_col, 9); 
				Print_cell_data(c_row, c_col, current, 9);
				current = current->left;
				c_col--;
				}
				
			}
			else if (c == 119) // up ke lia w
			{
				if (current->up == nullptr)
				{
					InsertRowAbove();
					Print_Grid();
					Print_Data();
					c_row++;
				}
				Print_cell(c_row, c_col, 9);
				Print_cell_data(c_row, c_col, current, 9);
				current = current->up;
				c_row--;
			}
			else if (c == 116) // t current ke upper dalne ke lia
			{
				if (current->up != nullptr)
				{
					InsertRowAbove();
					Print_Grid();
					Print_Data();
					c_row++;
				Print_cell(c_row, c_col, 9); 
				Print_cell_data(c_row, c_col, current, 9);
				current = current->up;
				c_row--;
				}
			}
			else if (c == 115) // down ke lia s
			{
				if (current->down == nullptr)
				{
					InsertRowBelow();
					Print_Grid();
					Print_Data();
				}
				Print_cell(c_row, c_col, 9); 
				Print_cell_data(c_row, c_col, current, 9);
				current = current->down;
				c_row++;
			}
			else if (c == 121) // y current ke neeche dale ga
			{
				if (current->down != nullptr)
				{
					InsertRowBelow();
					Print_Grid();
					Print_Data();

				}
				Print_cell(c_row, c_col, 9); 
				Print_cell_data(c_row, c_col, current, 9);
				current = current->down;
				c_row++;
			}
			else if (c == 105) // insertion ke lia i
			{
				insertion();
			}

			else if (c == 111) // shiff delete ke lia o
			{
				Shift_and_delete();
			}
			else if (c == 99) // Clear row and col ke lia c
			{
				ClearRow_Col();
			}

			else if (c == 109) // Mathematic operations ke lia m
			{
				Mathematic_operations(n1,n2,n3,n4);
			}
			else if (c == 120) // Cut and Copy ke lia x
			{
				Cut_Copy_Paste(n1,n2,n3,n4);
			}
			else if (c == 48) // Menu ke lia 0
			{
				menu();
			}
			gotoRowCol(c_row, c_col);
			Print_cell(c_row, c_col, 5); 
			Print_cell_data(c_row, c_col, current, 9);
		}
	}
void insertion()
{
    string input;

    while (true)
    {
        gotoRowCol(Cell_height * row_size + 10, 0);
        Print_cell_data(c_row, c_col, current, 9);
        gotoRowCol(Cell_height * row_size + 10, 0);
        cout << "Enter the value:" ;
        cin >> input;
		 gotoRowCol(Cell_height * row_size + 10, 0);
        cout << "                            " ;
        if (input.length() > 4 || !is_integer(input))
        {
			gotoRowCol(Cell_height * row_size + 10 , 0);
            cout << "Invalid input" << endl;
			system("pause");
            gotoRowCol(Cell_height * row_size + 10, 0);
            cout << "                                               ";
			gotoRowCol(Cell_height * row_size + 10+1, 0);
            cout << "                                               ";
			break;
        }
        else
        {
            gotoRowCol(Cell_height * row_size + 10+1, 0);
            cout << "                                                  ";
            gotoRowCol(Cell_height * row_size + 10 , 0);
            cout << "                                                 ";
            current->data = input;
            Print_cell_data(c_row, c_col, current, 9);
            break;
        }
   }
}
	void menu()
	{

		color(7);
		int x = 15; 
		gotoRowCol(0, col_size * x);
		cout << "Shortcut Keys: ";
		gotoRowCol(1, col_size * x);
		cout << "Insert : i";
		gotoRowCol(2, col_size * x);
		cout << "Shift Cell Right O then D";
		gotoRowCol(3, col_size * x);
		cout << "Shift Cell Down O then S";
		gotoRowCol(4, col_size * x);
		cout << "Delete Cell by Left Shift O then A";
		gotoRowCol(5, col_size * x);
		cout << "Delete Cell by Up Shift O then W";
		gotoRowCol(6, col_size * x);
		cout << "Delete Row O then R";
		gotoRowCol(7, col_size * x);
		cout << "Delete Col O then C";
		gotoRowCol(8, col_size * x);
		cout << "Clear Row C then C";
		gotoRowCol(9, col_size * x);
		cout << "Clear Col C then R";
		gotoRowCol(10, col_size * x);
		cout << "Copy X then C ";
		gotoRowCol(11, col_size * x);
		cout << "Cut X then X ";
		gotoRowCol(12, col_size * x);
		cout << "Paste X then P ";
		gotoRowCol(13, col_size * x);
		cout << "Perform Sum M then S ";
		gotoRowCol(14, col_size * x);
		cout << "Perform Average M then A ";
		gotoRowCol(15, col_size * x);
		cout << "Perform Count M then C ";
		gotoRowCol(16, col_size * x);
		cout << "Perform Max M then M ";
		gotoRowCol(17, col_size * x);
		cout << "Perform Min M then N ";
		gotoRowCol(18, col_size * x);
		cout << "Insert colomn on right of current e ";
		gotoRowCol(19, col_size * x);
		cout << "Insert colomn on left of current r ";
		gotoRowCol(20, col_size * x);
		cout << "Insert row on above of current t";
		gotoRowCol(21, col_size * x);
		cout << "Insert row on below of current y";

		gotoRowCol(22, col_size * x);
		system("pause");
		gotoRowCol(0, col_size * x);
		cout << "                                        ";
		gotoRowCol(1, col_size * x);
		cout << "                                        ";
		gotoRowCol(2, col_size * x);
		cout << "                                        ";
		gotoRowCol(3, col_size * x);
		cout << "                                        ";
		gotoRowCol(4, col_size * x);
		cout << "                                        ";
		gotoRowCol(5, col_size * x);
		cout << "                                        ";
		gotoRowCol(6, col_size * x);
		cout << "                                        ";
		gotoRowCol(7, col_size * x);
		cout << "                                        ";
		gotoRowCol(8, col_size * x);
		cout << "                                        ";
		gotoRowCol(9, col_size * x);
		cout << "                                        ";
		gotoRowCol(10, col_size * x);
		cout << "                                        ";
		gotoRowCol(11, col_size * x);
		cout << "                                        ";
		gotoRowCol(12, col_size * x);
		cout << "                                        ";
		gotoRowCol(13, col_size * x);
		cout << "                                        ";
		gotoRowCol(14, col_size * x);
		cout << "                                        ";
		gotoRowCol(15, col_size * x);
		cout << "                                        ";
		gotoRowCol(16, col_size * x);
		cout << "                                        ";
		gotoRowCol(17, col_size * x);
		cout << "                                        ";
		gotoRowCol(18, col_size * x);
		cout << "                                        ";
		gotoRowCol(19, col_size * x);
		cout << "                                        ";
		gotoRowCol(20, col_size * x);
		cout << "                                        ";
		gotoRowCol(21, col_size * x);
		cout << "                                        ";
		gotoRowCol(22, col_size * x);
		cout << "                                        ";
	}
	
};
int main()
{
	system("cls");
	Excel data;
	data.Keyboard();
}
