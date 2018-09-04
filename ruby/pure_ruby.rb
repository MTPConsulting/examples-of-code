require "json"
require "net/http"
require "uri"

module CurrencyConvert
  ##
  # Get currency value from ISO code currency and to ISO code currency
  #
  # @return [instance] Instanace of Client class
  #
  # @example
  #   client = CurrencyConvert::Client.new("EUR", "USD", 100)
  #   client.currency_value  #=> 1.14681
  #   client.current_from    #=> "EUR"
  #   client.current_to      #=> "USD"
  #   client.amount          #=> 100
  class Client
    attr_reader :currency_value
    attr_reader :currency_from
    attr_reader :currency_to
    attr_reader :currency_amount

    def initialize(currency_from, currency_to, amount)
      @currency_from = currency_from
      @currency_to = currency_to
      @amount = amount

      self.convert()
    end

    # Check if a currency_from and currency_to are valid
    def check_currency_valid
      url = "https://free.currencyconverterapi.com/api/v6/currencies"
      content = Net::HTTP.get(URI.parse(url))
      currencies = JSON.parse(content)["results"]

      if currencies[@currency_from] == nil
        raise "Currency #{@currency_from} not exists."
      end

      if currencies[@currency_to] == nil
        raise "Currency #{@currency_to} not exists."
      end

      true
    end

    # Convert currenty with google
    def convert
      if self.check_currency_valid()
        url = "https://free.currencyconverterapi.com/api/v6/convert?q=#{@currency_from}_#{@currency_to}&compact=y"
        content = Net::HTTP.get(URI.parse(url))
        begin
          @currency_value = JSON.parse(content)["#{@currency_from}_#{@currency_to}"]["val"]
          @currency_value = currency_value * @amount
        rescue NoMethodError
          @currency_value = 0
        end

        @currency_value
      else
        raise "Currency Error."
      end
    end
  end
end
