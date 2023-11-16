require "google_drive"

class GoogleDrive::Worksheet
    include Enumerable

    def method_missing(method_name, *args)
      ime = method_name.to_s
      indeks_kolone = kolona_indeks(ime)
      CelaKolona.new(self, indeks_kolone).sum if indeks_kolone
    end

    #def method_missing(method_name, *args)
     # ime = method_name.to_s
      #indeks_kolone = kolona_indeks(ime)
      #CelaKolona.new(self, indeks_kolone).avg if indeks_kolone
    #end

    def kolona_indeks(ime_kolone)
      indeks_kolone = rows.first.index { |name| name.downcase == ime_kolone.downcase }
      if indeks_kolone
        indeks_kolone += 1
        return indeks_kolone
      end
    end

    def each
      (1..num_rows).each do |indeks_reda|
        next if total_subtotal(indeks_reda)
          (1..num_cols).each do |indeks_kolone|
              next if self[indeks_reda, indeks_kolone] == ""
              yield self[indeks_reda, indeks_kolone].to_s
          end
      end
    end

  def row(indeks)
      vrednosti = []
      (1..num_cols).each {|indeks_kolone| vrednosti << self[indeks,indeks_kolone].to_s unless self[indeks,indeks_kolone].to_s.empty?}
      vrednosti
  end

  def total_subtotal(red_indeks)
      ima = (1..num_cols).find {|kolona_indeks| self[red_indeks,kolona_indeks].to_s == "total" || self[red_indeks,kolona_indeks].to_s == "subtotal"}
      ima
  end

  def spoji(drugi_ws)
    p "Nisu isti hederi" unless poklapanje_hedera(drugi_ws)
    
    (2..drugi_ws.num_rows).each do |i|
        red = drugi_ws.row(i)
        zamena_j = 1
        red.each do |j|
          prva = self[i,zamena_j].to_s
          druga = j.to_s
          spojeno = prva + druga
          self[i,zamena_j] = spojeno.to_s
          zamena_j = zamena_j + 1
        end
    end
    self.save
  end

  def razdvoji(drugi_ws)
      p "Nisu isti hederi" unless poklapanje_hedera(drugi_ws)

      (2..drugi_ws.num_rows).each do |i|
        red = drugi_ws.row(i)
        zamena_j = 1
        red.each do |j|
          prva = self[i,zamena_j].to_s
          druga = j.to_s
          razdvojeno = prva.gsub(druga,"")
          self[i,zamena_j] = razdvojeno.to_s
          zamena_j = zamena_j + 1
        end
    end
    self.save
  end

  def poklapanje_hedera(drugi_ws)
    self.rows[0] == drugi_ws.rows[0]
  end

    alias_method :originalna_get_implementacija, :[]
  
    def [](ime_kolone, *args)
      if ime_kolone.is_a?(String)
      indeks_kolone = rows.first.index { |name| name.downcase == ime_kolone.downcase }
  
        if indeks_kolone
        indeks_kolone += 1
  
    
            return CelaKolona.new(self, indeks_kolone)
          end
        end
      
  
      originalna_get_implementacija(ime_kolone, *args)
    
  end
   
end

class CelaKolona
  def initialize(worksheet, indeks_kolone)
    @worksheet = worksheet
    @indeks_kolone = indeks_kolone
  end

  def [](indeks)
    return @worksheet.column(@indeks_kolone) if indeks.nil?
    return @worksheet[indeks, @indeks_kolone]
  end

  def []=(indeks, vrednost)
    @worksheet[indeks, @indeks_kolone] = vrednost
    @worksheet.save
  end
  def to_s
      vrednosti = (1..@worksheet.num_rows).map { |i| @worksheet[i, @indeks_kolone].to_s }
      neprazni_stringovi = vrednosti.reject {|vrednost| vrednost.empty? }
      neprazni_stringovi.join(", ")  
  end

  def sum
    vrati = 0
    (2..@worksheet.num_rows).each do |i|
       vrati = vrati + @worksheet[i,@indeks_kolone].to_i
       p vrati
    end
  end

  def avg
    vrati = 0
    (2..@worksheet.num_rows).each do |i|
      vrati = vrati + @worksheet[i,@indeks_kolone].to_i
      p vrati
    end
  end

end


session = GoogleDrive::Session.from_config("config.json")

ws = session.spreadsheet_by_key("1nHLHHo88ExPDc_zvhP1-eVfn6NkSfZ9qQjkBJeXOPlY").worksheets[0]

ws2 = session.spreadsheet_by_key("1c6geVrAjoWfX0QQrZWpsIXck5vOTPtOZWGPRfwT4Twc").worksheets[0]

#p ws[1, 2] 


#1. Vraca tabelu u obliku matrice
#  p ws.rows

#2. Pristup jednom redu u tabeli i njegov ispis 
#    red = ws.row(3)
#    puts "Vrednosti u redu: #{red}"

#3. Ispisivanje cele mape pomocu each koji je implementiran u Google Drive klasi
#    ws2.each do |vrednost_iz_celije|
#      puts vrednost_iz_celije
#   end

#4. Biblioteka ne podrzava proveru da li je polje merge-ovano


#5.
#   cela_kolona = ws["Treca"]
#   puts "Cela kolona : #{cela_kolona}"

#   jedno_polje = ws["Cetvrta"][4]
#   p jedno_polje

#   ws["Prva"][3] = "test"

#8.     
#   ws.spoji(ws2)
#   ws.each do |vrednost|
 #    puts vrednost
 #  end
 
#9.
#    ws.razdvoji(ws2)
#    ws.each do |e|
#      puts e
#    end
#6.  
#    n = ws.Prva.sum

#   p n
  
