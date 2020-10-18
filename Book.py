class Book:
    def __init__(self, title, author, pages, series, genre, published, rating):
        self.title = title
        self.author = author
        self.series = series
        self.pages = pages
        self.genre = genre
        self.published = published
        self.rating = rating

    def __str__(self):
        temp = """Title: {0}
Author: {1}
Series: {2}
Pages: {3}
Genre: {4}
Published: {5}
GoodReads Rating: {6}
        """.format(self.title, self.author, self.series, self.pages, self.genre, self.published, self.rating)
        return temp
