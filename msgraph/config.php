<?php
require_once INCLUDE_DIR . 'class.plugin.php';
require_once 'rest.php';
class MSGraphPluginConfig extends PluginConfig {
    // Provide compatibility function for versions of osTicket prior to
    // translation support (v1.9.4)
    function translate() {
        if (!method_exists('Plugin', 'translate')) {
            return array(
                function($x) { return $x; },
                function($x, $y, $n) { return $n != 1 ? $y : $x; },
            );
        }
        return Plugin::translate('msigley-msgraph');
    }

    function getOptions() {
        list($__, $_N) = self::translate();
        return array(
            'msgraph-tenantId' => new TextboxField(array(
                'label' => $__('Tenant Id'),
                'configuration' => array('size'=>60, 'length'=>60)               
            )),
            'msgraph-clientId' => new TextboxField(array(
                'label' => $__('Client Id'),
                'configuration' => array('size'=>60, 'length'=>60)               
            )),
            'msgraph-clientSecret' => new PasswordField(array(
                'label' => $__('Client Secret'),
                'configuration' => array('size'=>60, 'length'=>60, 'placeholder' => str_repeat( '•', strlen( $this->get('msgraph-clientSecret') ) ) )               
            )),
            'msgraph-email' => new TextboxField(array(
                'label' => $__('Email'),
                'configuration' => array('size'=>60, 'length'=>60)                
            )),
            'msgraph-inboxFolder' => new TextboxField(array(
                'label' => $__('Name of Folder to Process New Messages from'),
                'configuration' => array('size'=>60, 'length'=>60)              
            )),
            'msgraph-processedFolder' => new TextboxField(array(
                'label' => $__('Name of Folder to Move Processed Messages to'),
                'configuration' => array('size'=>60, 'length'=>60)            
            )),
            'msgraph-emails-per-fetch' => new TextboxField(array(
                'label' => $__('Emails Per Fetch'),
                'configuration' => array('size'=>4, 'length'=>10)
            )),
            'msgraph-info' => new FreeTextField(array(
                'configuration' => array(
                    'html' => true, 
                    'size' => 'large',
                    'content' => '<strong>Emails are fetched as part of the cron task. Setting up external cron handling is highly recommended 
                        as documented <a href="https://docs.osticket.com/en/latest/Developer%20Documentation/API/Tasks.html?highlight=cron" target="_blank">here</a>.</strong>'
                ),
                'label' => ''
            ))
        );
    }

    /**
     * Pre-save hook to check configuration for errors (other than obvious
     * validation errors) prior to saving. Add an error to the errors list
     * or return boolean FALSE if the config commit should be aborted.
     */
    function pre_save(&$config, &$errors) {
        // Clear access token when config options are changed
        $this->set( 'msgraph-oauth-token_type', '' );
		$this->set( 'msgraph-oauth-access_token', '' );
		$this->set( 'msgraph-oauth-expires_in', 0 );

        // Clear cached folder ids
        $this->set( 'msgraph-inboxFolder-id', '' );
        $this->set( 'msgraph-processedFolder-id', '' );

        $config['msgraph-emails-per-fetch'] = (int) $config['msgraph-emails-per-fetch'];

        if( empty( $config['msgraph-tenantId'] ) || empty( $config['msgraph-clientId'] ) || ( empty( $config['msgraph-clientSecret'] ) && empty( $this->get( 'msgraph-clientSecret' ) ) )
            || !Validator::is_valid_email( $config['msgraph-email'] ) || empty( $config['msgraph-inboxFolder'] ) || empty( $config['msgraph-processedFolder'] ) 
            || $config['msgraph-emails-per-fetch'] < 1 || $config['msgraph-emails-per-fetch'] > 100 ) {
            $errors['err'] = 'Please correct the settings below.';
            return false;
        }

        // Test settings
        $rest = new MSGraphAPIREST( 
			array(
				'root_endpoint_url' => 'https://graph.microsoft.com/v1.0/', 
				'config' => &$this,
                'tenant_id' => $config['msgraph-tenantId'],
                'client_id' => $config['msgraph-clientId'],
                'client_secret' => $config['msgraph-clientSecret']
			)
		);

        $test = $rest->request( "users/", 'GET', array( 
            '$select' => "id" 
        ) );
        if( empty( $test ) ) {
            $errors['err'] = __( 'Unable to authenticate with the Microsoft Graph API. Please double check your Tenant Id, Client Id, and Client Secret settings.' );
            return false;
        }

        $test = $rest->request( "users/" . rawurlencode( $config['msgraph-email'] ), 'GET' );
        if( empty( $test ) ) {
            $errors['err'] = __( 'Unable to connect to the Email specified. Please double check your Email setting.' );
            return false;
        }


        $inboxFolderId = $rest->request( "users/" . rawurlencode( $config['msgraph-email'] ) ."/mailFolders", 'GET', array( 
            '$filter' => "displayName eq '" . $config['msgraph-inboxFolder'] . "'" 
        ) );
        
        if( $inboxFolderId && !empty( $inboxFolderId->value ) ) {
            $inboxFolderId = $inboxFolderId->value[0]->id;
            $this->set( 'msgraph-inboxFolder-id', $inboxFolderId );
        } else {
            $errors['err'] =  __( 'Unabled to find Process New Messages from folder.' );
            return false;
        }

        $processedFolderId = $rest->request( "users/" . rawurlencode( $config['msgraph-email'] ) ."/mailFolders", 'GET', array( 
            '$filter' => "displayName eq '" . $config['msgraph-processedFolder'] . "'" 
        ) );
        if( $processedFolderId && !empty( $processedFolderId->value ) ) {
            $processedFolderId = $processedFolderId->value[0]->id;
            $this->set( 'msgraph-processedFolder-id', $processedFolderId );
        } else {
            $errors['err'] =  __( 'Unabled to find Move Processed Messages to Folder.' );
            return false;
        }
        unset( $rest );

        return true;
    }
}
