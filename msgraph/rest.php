<?php
class MSGraphAPIREST {
private static $contruct_args = array( 'slug', 'root_endpoint_url', 'config', 'tenant_id', 'client_id', 'client_secret' );
	private static $flipped_construct_args = false;

	private $slug = __CLASS__;
	private $root_endpoint_url = 'https://example.com/v1/';
	private $config = null;

	private $tenant_id = '';
	private $client_id = '';
	private $client_secret = '';
	private $access_token = '';
	private $token_type = '';
	private $expires_in = 0;

	private $curl = null;
	private $curl_options = array(
		CURLOPT_FOLLOWLOCATION => true,
		CURLOPT_MAXREDIRS => 1,
		CURLOPT_REDIR_PROTOCOLS =>  CURLPROTO_HTTPS,
		CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
		// Disable TLS verification for speed
		CURLOPT_SSL_VERIFYHOST => 0,
		CURLOPT_SSL_VERIFYPEER => false,
		CURLOPT_TIMEOUT => 30,
		CURLOPT_CONNECTTIMEOUT_MS => 1000,
		CURLOPT_TCP_KEEPALIVE => 1,
		CURLOPT_DNS_USE_GLOBAL_CACHE => 1,
		CURLOPT_DNS_CACHE_TIMEOUT => 60 * 60, // 1 hour in seconds
		CURLOPT_RETURNTRANSFER => true,
		CURLINFO_HEADER_OUT => true
	);

	public function __construct( $args = array() ) {
		if( ! self::$flipped_construct_args ) {
			self::$contruct_args = array_flip( self::$contruct_args );
			self::$flipped_construct_args = true;
		}

		foreach( $args as $arg_name => $arg_value ) {
			if( empty($arg_value) || !isset( self::$contruct_args[$arg_name] ) )
				continue;

			if( !empty( self::$contruct_args[$arg_name] ) && defined( self::$contruct_args[$arg_name] ) )
				$this->$arg_name = constant( self::$contruct_args[$arg_name] );
			else
				$this->$arg_name = $arg_value;
		}

		if( $this->config !== null ) {
			if( empty( $this->tenant_id ) )
				$this->tenant_id = $this->config->get( 'msgraph-tenantId' );
			if( empty( $this->client_id ) )
				$this->client_id = $this->config->get( 'msgraph-clientId' );
			if( empty( $this->client_secret ) )
				$this->client_secret = $this->config->get( 'msgraph-clientSecret' );

			$this->token_type = (string) $this->config->get( 'msgraph-oauth-token_type' );

			$access_token = Crypto::decrypt( (string) $this->config->get( 'msgraph-oauth-access_token' ), SECRET_SALT, 'msgraph' );
			if( !empty( $access_token ) )
				$this->access_token = $access_token;
			
			$this->expires_in = (int) $this->config->get( 'msgraph-oauth-expires_in' );
		}
	}

	public function __destruct() {
		MSGraphAPIREST_shutdown( $this->curl );
    }

	private function get_authorization_header() {
		// Request new access token 2 minutes before it expires to prevent it from expiring mid request
		// Requests with messages with lots of attachments can take a while to complete
		if( $this->access_token === '' || $this->token_type === '' || $this->expires_in - 120 < time() ) {
			$oauth_rest = new MSGraphAPIREST( 
				array( 
					'root_endpoint_url' => 'https://login.microsoftonline.com/' . $this->tenant_id . '/oauth2/v2.0/' 
				) 
			);
			$oauth_response = $oauth_rest->request( 'token', 'POST', array(
				'client_id' => $this->client_id,
				'client_secret' => $this->client_secret,
				'scope' => 'https://graph.microsoft.com/.default',
				'grant_type' => 'client_credentials'
			), array(), true ); // Skip authorization on oauth request
			unset( $oauth_rest ); // Close connection

			$this->token_type = (string) $this->config->set( 'msgraph-oauth-token_type', $oauth_response->token_type );
			$this->access_token = (string) $oauth_response->access_token;
			$this->config->set( 'msgraph-oauth-access_token', Crypto::encrypt( $this->access_token, SECRET_SALT, 'msgraph' ) );
			$this->xpires_in = (int) $this->config->set( 'msgraph-oauth-expires_in', time() + (int) $oauth_response->expires_in );
		}

		return $this->token_type . ' ' . $this->access_token;
	}

	public function request( $endpoint, $method, $body = array(), $add_headers = array(), $skip_authorization = false, $output_file_pointer = false ) {
		if( !is_array( $body ) )
			return false;
		$endpoint = rtrim( ltrim( $endpoint, '/' ), '/' );
		$method = strtoupper( $method );
		$request_url = $this->root_endpoint_url . $endpoint;
		
		$headers = array( 
			'Accept' => 'application/json',
			'User-Agent' => '',
			'Prefer' => 'IdType="ImmutableId"' // https://learn.microsoft.com/en-us/graph/outlook-immutable-id
		) + $add_headers;

		if( $skip_authorization === false )
			$headers['Authorization'] = $this->get_authorization_header();
		
		if( $this->curl === null ) {
			$this->curl = curl_init();
			register_shutdown_function( 'MSGraphAPIREST_shutdown', $this->curl );
		}

		curl_reset( $this->curl );
		curl_setopt_array( $this->curl, $this->curl_options );

		$request_body = $body;
		$cache_bust = time();
		if( 'GET' == $method ) {
			curl_setopt( $this->curl, CURLOPT_HTTPGET, true );
			if( !empty( $request_body ) )
				$request_url .= '?' . http_build_query( $request_body, null, '&', PHP_QUERY_RFC3986 );
			$request_body = false;
		} elseif ( 'POST' == $method ) {
			curl_setopt( $this->curl, CURLOPT_POST, true );
			if( $headers['Content-Type'] === 'application/json' )
				$request_body = json_encode( $request_body );
		} else {
			curl_setopt( $this->curl, CURLOPT_CUSTOMREQUEST, $method );
			$headers['Content-Type'] = 'application/json';
			if( empty( $request_body ) )
				$request_body = false;
			else
				$request_body = json_encode( $request_body );
		}

		foreach( $headers as $name => &$value ) {
			$value = $name . ': ' . $value;
		}
		unset( $value );

		curl_setopt( $this->curl, CURLOPT_URL, $request_url );

		curl_setopt( $this->curl, CURLOPT_HTTPHEADER, $headers );
		if( false !== $request_body )
			curl_setopt( $this->curl, CURLOPT_POSTFIELDS, $request_body );

		if( $output_file_pointer !== false )
			curl_setopt( $this->curl, CURLOPT_FILE, $output_file_pointer );

		$response = curl_exec( $this->curl );
		
		// For debugging:
		// var_dump( curl_getinfo( $this->curl, CURLINFO_HEADER_OUT ) );
		// var_dump( $response );

		if( false === $response )
			return false;

		if( $output_file_pointer === false ) {
			$content_type = strtok( curl_getinfo( $this->curl, CURLINFO_CONTENT_TYPE ), ';' );
			if( $content_type === 'application/json' )
				$response = json_decode( $response );
		}

		if( !empty( $response->error ) )
			return false;

		return $response;
	}
}

function MSGraphAPIREST_shutdown( &$ch ) {
	if( $ch !== null )
		@curl_close( $ch );
	$ch = null;
}
