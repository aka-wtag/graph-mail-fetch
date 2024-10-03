package graph;

import java.util.Collections;
import java.util.List;

import com.azure.identity.ClientSecretCredential;
import com.azure.identity.ClientSecretCredentialBuilder;
import com.microsoft.graph.authentication.TokenCredentialAuthProvider;
import com.microsoft.graph.models.Message;
import com.microsoft.graph.requests.GraphServiceClient;
import com.microsoft.graph.requests.MessageCollectionPage;

public class GraphMailClient {

  public static void main(String[] args) {
    ClientSecretCredential clientSecretCredential = new ClientSecretCredentialBuilder().clientId(
            "clientId")
        .clientSecret("clientSecret")
        .tenantId("tenantId")
        .build();

    TokenCredentialAuthProvider authProvider = new TokenCredentialAuthProvider(
        Collections.singletonList("https://graph.microsoft.com/.default"), clientSecretCredential);

    GraphServiceClient<?> graphClient = GraphServiceClient.builder()
        .authenticationProvider(authProvider)
        .buildClient();

    // Fetch mails
    fetchMails(graphClient);
//    fetchUnreadMails(graphClient);
  }

  // Fetch mails in inbox
  private static void fetchMails(GraphServiceClient<?> graphClient) {
    System.out.println("Fetching mails...");
    MessageCollectionPage messages = graphClient.users("id")
        .messages()
        .buildRequest()
        .select("subject,from,receivedDateTime") // Select specific fields
        .top(10) // Get the top 10 emails
        .get();
    System.out.println("Fetching mails done...");

    // print the messages
    List<Message> mailList = messages.getCurrentPage();
    for (Message mail : mailList) {
      System.out.println("Subject: " + mail.subject);
      System.out.println("From: " + mail.from.emailAddress.address);
      System.out.println("Received: " + mail.receivedDateTime);
      System.out.println("-------------------");
    }

  }

  // Fetch the unread mails in inbox
  private static void fetchUnreadMails(GraphServiceClient<?> graphClient) {
    System.out.println("Fetching mails...");
    MessageCollectionPage unreadMessages = graphClient.users("devtesting@welldev.io")
        .messages()
        .buildRequest()
        .filter("isRead eq false") // Filter to only get unread emails
        .select("subject,from,receivedDateTime,isRead") // Select specific fields
        .top(10) // Get the top 10 unread emails
        .get();
    System.out.println("Fetching mails done...");

    // print the unread messages
    List<Message> unreadMailList = unreadMessages.getCurrentPage();
    for (Message mail : unreadMailList) {
      System.out.println("Subject: " + mail.subject);
      System.out.println("From: " + mail.from.emailAddress.address);
      System.out.println("Received: " + mail.receivedDateTime);
      System.out.println("-------------------");

      markEmailAsRead(graphClient, mail.id);
    }
  }

  // Mark the mail as read
  private static void markEmailAsRead(GraphServiceClient<?> graphClient, String messageId) {
    graphClient
        .users("devtesting@welldev.io")
        .messages(messageId)
        .buildRequest()
        .patch(new Message() {
          {
            isRead = true; // Set isRead to true
          }
        });

    System.out.println("Message with ID " + messageId + " marked as read.");
  }
}
