def num2col(num):
	"""Excel-style columns are designated by letters, A, B, ..., Z, AA, AB, ..., ZZ, AAA, etc. This converts an int to that format"""
	col = ""
	while num > 0:
		num -= 1
		col = chr(ord("a") + num % 26) + col
		num /= 26
	return col

def col2num(col):
	"""Convert Excel-style column name to a number suitable for use as an array index"""
	num = 0
	for c in col:
		num *= 26
		num += (ord(c.upper()) - ord('A')) + 1
	return num

class Sheet(object):
	def __init__(self):
		self.width = 4	# Stored length of each row. Length of columns doesn't require any effort to keep in check.
		self.cells = []	# The formulaic value of the cells
		self.state = []	# The most recent computed value of the cells
	def dump(self):
		"""Print the whole sheet's internal values"""
		print "Cell contents:"
		for row in self.cells:
			print row
		print "Cell values:"
		for row in self.state:
			print row
	def set(self, col, row, value):
		"""Set a cell's value given column and row"""
		if col >= self.width:
			for r in self.cells + self.state:
				r.extend([0] * (col + 1 - self.width))
			self.width = col + 1
		if row >= len(self.cells):
			self.cells.extend([[0] * self.width for n in xrange(row - len(self.cells) + 1)])
		self.cells[row][col] = value
	def get(self, col, row):
		"""Fetch a value given column and row"""
		if row < len(self.cells):
			if col < self.width:
				if type(self.cells[row][col]) in (int, long, float, complex):
					return self.cells[row][col]
				if row < len(self.state):
					return self.state[row][col]
		return 0
	def dict(self):
		out = {}
		for rownum, row in enumerate(self.state):
			for colnum, cell in enumerate(row):
				out[num2col(colnum + 1) + str(rownum)] = self.get(colnum, rownum)
		return out
	def eval(self):
		"""Evaluate any formulae found in the sheet (string values in cells)"""
		self.state = [x[:] for x in self.cells]
		for rownum, row in enumerate(self.cells):
			for colnum, cell in enumerate(row):
				if type(cell) is str:
					try:
						self.state[rownum][colnum] = eval(cell, self.dict());
					except:
						pass

if __name__ == "__main__":
	sheet = Sheet()
	sheet.set(0, 0, 3)
	sheet.set(1, 0, 4)
	sheet.set(2, 3, 42)
	sheet.set(6, 6, 18)
	sheet.set(3, 3, "a0 + b0")
	sheet.eval()
	print sheet.dict()
	sheet.dump()
