<?php
require_once 'config.php';
require_once 'rest.php';
require_once INCLUDE_DIR.'api.tickets.php';
require_once INCLUDE_DIR . "class.mailparse.php";
class MSGraphPlugin extends Plugin {
	var $config_class = "MSGraphPluginConfig";

	private $ost = null;
	private $config = null;
	private $rest = null;

	function bootstrap() {
		global $ost;

		$this->ost = &$ost;
		$this->config = $this->getConfig();
		Signal::connect( 'cron', array($this, 'cron') );
	}

	function cron() {
		set_time_limit( 30 * (int) $this->config->get( 'msgraph-emails-per-fetch' ) );

		// API Limits
		// https://learn.microsoft.com/en-us/graph/throttling-limits#outlook-service-limits
		// As of 09/23/2022
		// 10,000 API requests in a 10 minute period
		// 4 concurrent requests

		// Authorization and OAuth tokens are handled in MSGraphAPIREST
		$this->rest = new MSGraphAPIREST( 
			array(
				'root_endpoint_url' => 'https://graph.microsoft.com/v1.0/', 
				'config' => &$this->config
			) 
		);

		$inboxFolderId = $this->config->get( 'msgraph-inboxFolder-id' );
		if( empty( $inboxFolderId ) ) {
			$inboxFolderId = $this->rest->request( "users/" . rawurlencode( $this->config->get( 'msgraph-email' ) ) ."/mailFolders", 'GET', array( 
				'$filter' => "displayName eq '" . $this->config->get( 'msgraph-inboxFolder' ) . "'" 
			) );
			if( $inboxFolderId && !empty( $inboxFolderId->value ) ) {
				$inboxFolderId = $inboxFolderId->value[0]->id;
				$this->config->set( 'msgraph-inboxFolder-id', $inboxFolderId );
			} else {
				$this->ost->logWarning( __( 'Microsoft Graph Error' ), __( 'Unabled to find Process New Messages from folder.' ), false );
				unset( $this->rest ); // Close connection
				return;
			}
		}

		$processedFolderId = $this->config->get( 'msgraph-processedFolder-id' );
		if( empty( $processedFolderId ) ) {
			$processedFolderId = $this->rest->request( "users/" . rawurlencode( $this->config->get( 'msgraph-email' ) ) ."/mailFolders", 'GET', array( 
				'$filter' => "displayName eq '" . $this->config->get( 'msgraph-processedFolder' ) . "'" 
			) );
			if( $processedFolderId && !empty( $processedFolderId->value ) ) {
				$processedFolderId = $processedFolderId->value[0]->id;
				$this->config->set( 'msgraph-processedFolder-id', $processedFolderId );
			} else {
				$this->ost->logWarning( __( 'Microsoft Graph Error' ), __( 'Unabled to find Move Processed Messages to Folder.' ), false );
				unset( $this->rest ); // Close connection
				return;
			}
		}

		// Filter is used to remove Teams messages: https://learn.microsoft.com/en-us/graph/known-issues#get-messages-returns-chats-in-microsoft-teams
		// Filter and orderby quirk: https://learn.microsoft.com/en-us/graph/api/user-list-messages?view=graph-rest-1.0&tabs=http#using-filter-and-orderby-in-the-same-query
		$messages = $this->rest->request( "users/" . rawurlencode( $this->config->get( 'msgraph-email' ) ) . "/mailFolders/" . $inboxFolderId . "/messages", 'GET', array( 
			'$select' => 'id,subject,receivedDateTime', '$filter' => "receivedDateTime ge 1970-01-01 and subject ne 'IM'", '$orderby' => 'receivedDateTime asc',
			'$top' => $this->config->get( 'msgraph-emails-per-fetch' ) ) );

		if( !empty( $messages ) )
			$messages = $messages->value;

		if( empty( $messages ) ) {
			// No new messages to process
			unset( $this->rest ); // Close connection
			return;
		}
		
		// Process messages into OSTicket
		foreach( $messages as &$message ) {
			$this->processMessage( $message );
		}

		// Move processed messages
		$move_requests = array();
		$move_request_id = 1;
		foreach( $messages as $message ) {
			$move_request = new StdClass;
			$move_request->id = $move_request_id;
			// Each request in batch depends on the previous one to prevent MailboxConcurrency limit issues
			if( $move_request_id > 1 )
				$move_request->dependsOn = array( $move_request_id - 1 );
			$move_request->method = 'POST';
			// POST /users/{id | userPrincipalName}/mailFolders/{id}/messages/{id}/move
			$move_request->url = "/users/" . rawurlencode( $this->config->get( 'msgraph-email' ) ) . "/messages/" . $message->id . '/move';
			$move_request->body = new StdClass;
			$move_request->body->destinationId = $processedFolderId;
			$move_request->headers = array( 'Content-Type' => 'application/json' );
			$move_requests[] = $move_request;

			$move_request_id++;
		}

		$move_response = $this->rest->request( '$batch', 'POST', array( 'requests' => $move_requests ), array( 'Content-Type' => 'application/json' ) );

		unset( $this->rest ); // Close connection
	}

	// Based on logic from MailFetcher::createTicket()
	private function processMessage( $message ) {
		$tempfile = tempnam( sys_get_temp_dir(), 'OME' ); // Window only supports three character temp filename prefixes
		$fp = fopen( $tempfile, 'w+' );
		
		// TODO: Save raw messages to files and process them in a second pass to better handle timeouts
		$response = $this->rest->request( "users/" . rawurlencode( $this->config->get( 'msgraph-email' ) ) . "/messages/" . $message->id . '/$value', 'GET', array(), array(), false, $fp );

		if( !empty( $response ) ) {
			$api = new MSGraphTicketApiController();
			$parser = new ApiEmailDataParser();
			fseek( $fp, 0, SEEK_SET );
			if( $data = $parser->parse( $fp ) ) {
				fclose( $fp );
				unlink( $tempfile );
				return $api->processEmail( $data );
			}
		}

		fclose( $fp );
		unlink( $tempfile );
		return false;
	}
}

class MSGraphTicketApiController extends TicketApiController {
	/**
	 * API error & logging and response!
	 *
	 */

	/* Overridden to prevent exit calls in response() to allow tickets to be denied and emails to continue to be processed */
	function exerr($code, $error='') {
		global $ost;

		if($error && is_array($error))
			$error = Format::array_implode(": ", "\n", $error);

		//Log the error as a warning - include api key if available.
		$msg = $error;
		if($_SERVER['HTTP_X_API_KEY'])
			$msg.="\n*[".$_SERVER['HTTP_X_API_KEY']."]*\n";
		$ost->logWarning(__('API Error')." ($code)", $msg, false);

		fwrite(STDERR, "({$code}) $error\n");

		return false;
	}
}