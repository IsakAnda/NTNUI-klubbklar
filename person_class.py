import pandas as pd

template = 'mal.xlsx'                               # Excel file for the diffrent product names and descriptions
prices = 'prisliste.xlsx'                           # Excel file for the price of the diffrent products
coaches = 'trenerrabatt_2627.xlsx'                  # Excel file for a list of the coaches and how much discount they should get.

template_df = pd.read_excel(template)               
price_df = pd.read_excel(prices)                    
coaches_df = pd.read_excel(coaches, skiprows=2)     

class Person:
    def __init__(self,row):
        self.name = row['Your full name']
        self.team = row['Which team are you playing for? ']
        self.number = row['Your phone number']
        self.mail = row['E-postadresse']
        self.timestamp = row['Tidsmerke']
        self.model = 'herre' if row['Which model do you want?'] == 'Men' else 'dame'

        self.is_coach = row['Do you have a coaching position?'] != 'Nei/ No'
        self.coach_offer =0
        self.used_coach = (False, None)
        self.coach_product = None
        self.coach_error = None

        if self.name in coaches_df.get('Navn').to_list():
            self.coach_offer = coaches_df['Trenerrabatt i [kr]'].to_list()[coaches_df.get('Navn').to_list().index(self.name)]
            self.used_coach = (False, None)     # FIXME handel already ordered products.
            self.coach_product = None
        elif self.is_coach and self.name not in coaches_df.get('Navn').to_list():
            print('here there is somthing wrong with coach discount!!:', self.name)
            self.coach_error = f'{self.name} er ikke trener?'
        
        self.row = row
        self.order = []
        self.invoice = []
        self.total_price = 0

        self.create_order()

    def addPrice(self, product, name_print=None, coachDiscount=False, comment = None):
        if name_print == None:
            price = int(price_df.loc[price_df[price_df.columns[0]] == product, price_df.columns[1]].values[0])
        else:
            price = int(price_df.loc[price_df[price_df.columns[0]] == product, price_df.columns[2]].values[0])
        price = price-self.coach_offer if coachDiscount else price
        if coachDiscount:
            if self.coach_error != None:
                if comment != None:
                    comment+=self.coach_error
                else:
                    comment=self.coach_error
            elif not self.used_coach[0]:
                self.used_coach=(True, product)
            

        new_dict = {
            'Lag': self.team,
            'Navn': self.name,
            'Produkt': product,
            'Trykk?': name_print,
            'Trenerrabatt': self.coach_offer if coachDiscount else None,
            'Kostnad enkelt produkt': price,
            'Kostnad enkelt personer': price,
            'Totalt per lag': price,
            'Kommmentar': comment
        }
        if coachDiscount and self.coach_offer >0:
            self.used_coach = (True, self.timestamp)
            self.coach_product = product
            self.coach_offer=0
        self.invoice.append(new_dict.copy())
        self.total_price+=price

    def addItem(self, product, size, model = 'unisex', color='Sort', name_print = None, comment = None):
        result = template_df[template_df.iloc[:, 0] == product].iloc[0]
        new_dict = {
            'Lag': self.team,
            'Navn': self.name,
            'Produkt': product,
            'Merke': result['Merke'],
            'Modellnavn': result['Modellnavn'],
            'Modell': model,
            'Farge': color,
            'Størrelse': size,
            'Navnetrykk (rygg)': name_print,
            'Kommentar': comment}
        self.order.append(new_dict.copy())

    def addTshirt(self):
        if self.row['Size of t-shirt(s)'] == 'Skal ikke ha':
            return
        size = self.row['Size of t-shirt(s)']
        black = int(self.row['Amount of black t-shirts'])
        red = int(self.row['Amount of red t-shirts'])
        white = int(self.row['Amount of white t-shirts'])
        name_print = self.row['If yes, what should be printed on the back?'] if self.row['Do you want a name print?'] == 'Ja/ Yes' else None
        comment = self.row['Comments regarding the t-shirt order.\n'] if self.row['Comments regarding the t-shirt order.\n'] != 0 else None

        colors = ['Sort']*black+['Rød']*red+['Hvit']*white
        for color in colors:
            self.addItem(f'Tskjorte ({self.model})', size, model=self.model, color=color, name_print=name_print, comment=comment)
            self.addPrice('Tskjorte', name_print=name_print, comment=comment)

    def addWarmUpSweatshirt(self):
        if self.row['Size of warm-up sweatshirt(s)\n'] == 'Skal ikke ha':
            return
        size = self.row['Size of warm-up sweatshirt(s)\n']
        black = int(self.row['Amount of black sweatshirts'])
        green = int(self.row['Amount of green sweatshirts'])
        red = int(self.row['Amount of red sweatshirts'])
        name_print = self.row['If yes, what should be printed on the back? 2'] if self.row['Do you want name print?'] == 'Ja/ Yes' else None
        comment = self.row['Comments regarding the sweatshirts'] if self.row['Comments regarding the sweatshirts'] != 0 else None

        colors = ['Sort']*black+['Grønn']*green+['Rød']*red

        for color in colors:
            self.addItem(f'Oppvarmingsgenser ({self.model})', size, model=self.model, color=color, name_print=name_print, comment=comment)
            self.addPrice('Oppvarmingsgenser', name_print=name_print, comment=comment)

    def addkHalfZip(self):
        amount = int(self.row['Amount of half-zips'])
        size = self.row['Size of half-zip']
        name_print = self.row['If yes, what should be printed on the back? 3'] if self.row['Do you want name print? 2'] == 'Ja/ Yes' else None

        for _ in range(amount):
            self.addItem(f'Half-zip ({self.model})', size, model=self.model, name_print=name_print)
            self.addPrice('Half-zip', name_print=name_print)

    def addHoodie(self):
        amount = int(self.row['Amount of hoodies?'])
        size = self.row['Size of hoodie']
        name_print = None if self.row['Name to be printed on the back']==0 else self.row['Name to be printed on the back']

        for _ in range(amount):
            self.addItem(f'Hettegenser ({self.model})', size, model=self.model, name_print=name_print)
            self.addPrice('Hettegenser', name_print=name_print)

    def addHoodjacket(self):
        amount = int(self.row['Amount of hoodjackets'])
        size = self.row['Size of hoodjacket']
        name_print = None if self.row['Name to be printed on the back 2']==0 else self.row['Name to be printed on the back 2']
        coachdeal = self.row['Trenerrabatt'] != 'Nei/ No'

        comment = 'Trener %' if coachdeal else None

        for _ in range(amount):
            self.addItem(f'Hettejakke ({self.model})', size, model=self.model, name_print=name_print, comment=comment)
            self.addPrice('Hettejakke', name_print=name_print, coachDiscount=coachdeal)


    def addCrewneck(self):
        amount = int(self.row['Amount of crewnecks'])
        size = self.row['Size of crewneck']
        name_print = None if self.row['Name to be printed on the back 3']==0 else self.row['Name to be printed on the back 3']

        for _ in range(amount):
            self.addItem(f'Crewneck (unisex)', size, name_print=name_print)
            self.addPrice('Crewneck', name_print=name_print)

    def addwomensShorts(self):
        amount = int(self.row["Amount of women's shorts"])
        size = self.row["Size women's shorts"]

        for _ in range(amount):
            self.addItem(f'Shorts (dame)', size, 'dame')
            self.addPrice('Shorts (dame)')

    def addmensShorts(self):
        amount = int(self.row["Amount of men's shorts"])
        size = self.row["Size men's shorts"]

        for _ in range(amount):
            self.addItem(f'Shorts (herre)', size, 'herre')
            self.addPrice('Shorts (herre)')

    def addsweatpants(self):
        amount = int(self.row["Amount of sweatpants"])
        size = self.row["Size sweatpants"]

        for _ in range(amount):
            self.addItem(f'Joggebukse (unisex)', size)
            self.addPrice('Joggebukse')

    def addEverydaySweaters(self):
        white = int(self.row['Amount of white everyday sweaters'])
        black = int(self.row['Amount of black everyday sweaters'])
        blue = int(self.row['Amount of navy blue everyday sweaters'])
        size = self.row['Size of everyday sweaters']

        colors = ['Hvit']*white+['Sort']*black+['Marineblå']*blue
        for color in colors:
            self.addItem('Tynn genser (bomull)', size, color)
            self.addPrice('Tynn genser (bomull)')

    def addBag(self):
        bag38 = int(self.row['Amount of bags (38L)'])
        bag65 = int(self.row['Amount of backpacks (65L)'])

        for _ in range(bag65):
            self.addItem('Bag (stor)', 'L')
            self.addPrice('Bag (65 L)')
        for _ in range(bag38):
            self.addItem('Bag (liten)', 'S')
            self.addPrice('Bag (38 L)')

    def addExtras(self):
        sleeves = int(self.row['Amount of sleeves'])
        for _ in range(sleeves):
            self.addItem('Sleeves', None)
            self.addPrice('Sleeves')
        
        amount_short = int(self.row['Amount of knee pads (short)'])
        size_ws = self.row['Size white knee pads (short)']
        size_bs = self.row['Size black knee pads (short)']
        colors = []
        sizes = []

        if size_ws != 'Skal ikke ha/ None': 
            colors= ['Hvit']*amount_short
            sizes=[size_ws]*amount_short
        if size_bs != 'Skal ikke ha/ None':
            if colors == []:
                colors = ['Sort']*amount_short
                sizes = [size_bs]*amount_short
            else:
                colors[-amount_short//2:]=['Sort']*(amount_short//2)
                sizes[-amount_short//2:] = [size_bs]*(amount_short//2)

        for i in range(len(sizes)):
            self.addItem('Knebeskytter (kort)', sizes[i], color=colors[i])
            self.addPrice('Knebeskytter')

        amount_long = int(self.row['Amount knee pads (long)'])
        size_wl = self.row['Size white knee pads (long)']
        size_bl = self.row['Size black knee pads (long)']
        colors = []
        sizes = []

        if size_wl != 'Skal ikke ha/ None': 
            colors= ['Hvit']*amount_long
            sizes=[size_wl]*amount_long
        if size_bl != 'Skal ikke ha/ None':
            if colors == []:
                colors = ['Sort']*amount_long
                sizes = [size_bl]*amount_long
            else:
                colors[-amount_long//2:]=['Sort']*(amount_long//2)
                sizes[-amount_long//2:] = [size_bl]*(amount_long//2)

        for i in range(len(sizes)):
            self.addItem('Knebeskytter (lang)', sizes[i], color=colors[i])
            self.addPrice('Knebeskytter')



    def addCompressionShirt(self):
        size = self.row['Size on compression shirts']
        black = int(self.row['Amount of black compression shirts'])
        red = int(self.row['Amount of red compression shirts'])
        green = int(self.row['Amount of green compression shirts'])

        colors = ['Sort']*black+['Rød']*red+['Grønn']*green
        for color in colors:
            self.addItem('Kompresjonsgenser (unisex)', size, color=color)
            self.addPrice('Kompresjonsgenser')

    def create_order(self):
        self.addTshirt()
        self.addWarmUpSweatshirt()
        self.addkHalfZip()
        self.addHoodie()
        self.addHoodjacket()
        self.addCrewneck()
        self.addwomensShorts()
        self.addmensShorts()
        self.addsweatpants()
        self.addEverydaySweaters()
        self.addBag()
        self.addExtras()
        self.addCompressionShirt()

        for row in self.invoice:
            row['Kostnad enkelt personer']=self.total_price
