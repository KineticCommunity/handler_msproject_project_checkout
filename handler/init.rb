# Require the dependencies file to load the vendor libraries
require File.expand_path(File.join(File.dirname(__FILE__), 'dependencies'))

class MsprojectProjectCheckoutV1
  def initialize(input)
    # Set the input document attribute
    @input_document = REXML::Document.new(input)

    # Store the info values in a Hash of info names to values.
    @info_values = {}
    REXML::XPath.each(@input_document,"/handler/infos/info") { |item|
      @info_values[item.attributes['name']] = item.text
    }
    @enable_debug_logging = @info_values['enable_debug_logging'] == 'Yes'

    # Store parameters values in a Hash of parameter names to values.
    @parameters = {}
    REXML::XPath.match(@input_document, '/handler/parameters/parameter').each do |node|
      @parameters[node.attribute('name').value] = node.text.to_s
    end
  end

  def execute()
    resources_path = File.join(File.expand_path(File.dirname(__FILE__)), 'resources')

    # Create the command string that will be used to retrieve the cookies
    cmd_string = "O365Auth.Console.exe #{@info_values['ms_project_location']} #{@info_values['username']} #{@info_values['password']} #{@info_values['integrated_authentication']}"

    # Retrieve the cookies
    cookies = `cd "#{resources_path}" & #{cmd_string}`

    proj_resource = RestClient::Resource.new(@info_values['ms_project_location'].chomp("/"),
      :headers => {:content_type => "application/json",:accept => "application/json", :cookie => cookies})
    
    context_endpoint = proj_resource["/_api/contextinfo"]
    puts "Sending a request to get the FormDigestValue that will be passed at the X-RequestDigest header in the create call" if @enable_debug_logging
    begin
      results = context_endpoint.post ""
    rescue RestClient::Exception => error
      raise StandardError, error.inspect
    end

    json = JSON.parse(results)
    # Get the JSON value array that contains the lookup table information
    form_digest_value = json["FormDigestValue"]
    proj_resource.headers["X-RequestDigest"] = form_digest_value

    checkout_endpoint = proj_resource["/_api/ProjectServer/Projects('#{@parameters['project_id']}')/checkOut()/IncludeCustomFields"]

    retry_num = 0
    need_retry = true
    error_copy = nil
    while need_retry == true && retry_num < 12
      begin
        puts "Checking out project '#{@parameters['project_id']}'" if @enable_debug_logging
        checkout_endpoint.post ""
        need_retry = false
      rescue RestClient::Exception => error
        if error.http_code == 403
          puts "Server returned a non fatal 403 forbidden. Attempting again #{11-retry_num} more time(s)" if @enable_debug_logging
          retry_num += 1
          error_copy = error
          sleep(10)
        else
          raise StandardError, error.inspect
        end
      end
    end

    if need_retry == true
      raise StandardError, error_copy.inspect
    end

    # Return the results
    <<-RESULTS
    <results/>
    RESULTS
  end

  # This is a template method that is used to escape results values (returned in
  # execute) that would cause the XML to be invalid.  This method is not
  # necessary if values do not contain character that have special meaning in
  # XML (&, ", <, and >), however it is a good practice to use it for all return
  # variable results in case the value could include one of those characters in
  # the future.  This method can be copied and reused between handlers.
  def escape(string)
    # Globally replace characters based on the ESCAPE_CHARACTERS constant
    string.to_s.gsub(/[&"><]/) { |special| ESCAPE_CHARACTERS[special] } if string
  end
  # This is a ruby constant that is used by the escape method
  ESCAPE_CHARACTERS = {'&'=>'&amp;', '>'=>'&gt;', '<'=>'&lt;', '"' => '&quot;'}
end